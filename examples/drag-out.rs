// Copyright 2014-2021 The winit contributors
// Copyright 2021-2023 Tauri Programme within The Commons Conservancy
// SPDX-License-Identifier: Apache-2.0

use std::{cell::RefCell, ffi::OsString, mem::ManuallyDrop, os::windows::prelude::OsStrExt};

use tao::{
  event::{ElementState, Event, MouseButton, WindowEvent},
  event_loop::{ControlFlow, EventLoop},
  window::WindowBuilder,
};
use windows::{
  core::HRESULT,
  Win32::{
    Foundation::{
      GlobalFree, BOOL, DRAGDROP_S_CANCEL, DRAGDROP_S_DROP, DV_E_FORMATETC, E_NOTIMPL,
      E_OUTOFMEMORY, HGLOBAL, OLE_E_ADVISENOTSUPPORTED, S_OK,
    },
    System::{
      Com::{
        IAdviseSink, IBindCtx, IDataObject, IDataObject_Impl, IEnumFORMATETC, IEnumSTATDATA,
        DATADIR_GET, DVASPECT_CONTENT, FORMATETC, STGMEDIUM, TYMED_HGLOBAL,
      },
      Memory::{GlobalAlloc, GlobalLock, GlobalSize, GlobalUnlock, GLOBAL_ALLOC_FLAGS, GMEM_FIXED},
      Ole::{
        DoDragDrop, IDropSource, IDropSource_Impl, CF_TEXT, DROPEFFECT_COPY, DROPEFFECT_LINK,
        DROPEFFECT_NONE,
      },
      SystemServices::{MK_LBUTTON, MK_RBUTTON, MODIFIERKEYS_FLAGS},
    },
    UI::Shell::{
      IDataObjectAsyncCapability, IDataObjectAsyncCapability_Impl, SHCreateStdEnumFmtEtc,
    },
  },
};
use windows_implement::implement;

const DATA_E_FORMATETC: HRESULT = HRESULT(-2147221404 + 1);

#[implement(IDropSource)]
struct DragDropClient {}

#[allow(non_snake_case)]
impl IDropSource_Impl for DragDropClient {
  fn QueryContinueDrag(
    &self,
    fescapepressed: BOOL,
    grfkeystate: MODIFIERKEYS_FLAGS,
  ) -> ::windows::core::HRESULT {
    if fescapepressed.as_bool() {
      return DRAGDROP_S_CANCEL;
    }

    if (grfkeystate & (MK_LBUTTON | MK_RBUTTON)).0 == 0 {
      return DRAGDROP_S_DROP;
    }

    return S_OK;
  }

  fn GiveFeedback(
    &self,
    _dweffect: windows::Win32::System::Ole::DROPEFFECT,
  ) -> ::windows::core::HRESULT {
    return windows::Win32::Foundation::DRAGDROP_S_USEDEFAULTCURSORS;
  }
}

fn duplicate_global_data(global: HGLOBAL) -> windows::core::Result<HGLOBAL> {
  unsafe {
    let len = GlobalSize(global);
    let src = GlobalLock(global);
    let dest = GlobalAlloc(GMEM_FIXED, len)?;
    std::ptr::copy_nonoverlapping(src, dest.0 as _, len);
    let _ = GlobalUnlock(global);
    Ok(dest)
  }
}

fn global_from_data(data: &[u16]) -> windows::core::Result<HGLOBAL> {
  unsafe {
    let global = GlobalAlloc(GLOBAL_ALLOC_FLAGS(0), data.len())?;
    let global_data = GlobalLock(global);
    if global_data.is_null() {
      GlobalFree(global)?;
      Err(E_OUTOFMEMORY.into())
    } else {
      std::ptr::copy_nonoverlapping(data.as_ptr(), global_data as *mut u16, data.len());
      let _ = GlobalUnlock(global);
      Ok(global)
    }
  }
}

#[implement(IDataObject, IDataObjectAsyncCapability)]
struct DragDropObject {
  fmtetc: Vec<FORMATETC>,
  stgmeds: Vec<STGMEDIUM>,
  fdoopasync: RefCell<bool>,
  inoperation: RefCell<bool>,
}

impl DragDropObject {
  fn lookup_format(&self, pformatetc: *const FORMATETC) -> Option<usize> {
    let format = unsafe { *pformatetc };
    self.fmtetc.iter().position(|e| {
      e.cfFormat == format.cfFormat
        && (e.tymed & format.tymed) != 0
        && e.dwAspect == format.dwAspect
        && e.lindex == format.lindex
    })
  }
}

#[allow(non_snake_case)]
impl IDataObject_Impl for DragDropObject {
  fn GetData(&self, pformatetcin: *const FORMATETC) -> ::windows::core::Result<STGMEDIUM> {
    match self.lookup_format(pformatetcin) {
      None => Err(DV_E_FORMATETC.into()),
      Some(idx) => {
        let mut stgmed = STGMEDIUM::default();
        stgmed.tymed = self.fmtetc[idx].tymed;
        stgmed.pUnkForRelease = ManuallyDrop::new(None);
        if self.fmtetc[idx].tymed as i32 == TYMED_HGLOBAL.0 {
          stgmed.u.hGlobal = duplicate_global_data(unsafe { self.stgmeds[idx].u.hGlobal })?
        }
        Ok(stgmed)
      }
    }
  }

  fn GetDataHere(
    &self,
    _pformatetc: *const FORMATETC,
    _pmedium: *mut STGMEDIUM,
  ) -> ::windows::core::Result<()> {
    Err(DATA_E_FORMATETC.into())
  }

  fn QueryGetData(&self, pformatetc: *const FORMATETC) -> ::windows::core::HRESULT {
    self
      .lookup_format(pformatetc)
      .map(|_| S_OK)
      .unwrap_or(DV_E_FORMATETC)
  }

  fn GetCanonicalFormatEtc(
    &self,
    _pformatectin: *const FORMATETC,
    pformatetcout: *mut FORMATETC,
  ) -> ::windows::core::HRESULT {
    unsafe {
      (*pformatetcout).ptd = std::ptr::null_mut();
    }
    E_NOTIMPL
  }

  fn SetData(
    &self,
    _pformatetc: *const FORMATETC,
    _pmedium: *const STGMEDIUM,
    _frelease: BOOL,
  ) -> ::windows::core::Result<()> {
    Err(E_NOTIMPL.into())
  }

  fn EnumFormatEtc(&self, dwdirection: u32) -> ::windows::core::Result<IEnumFORMATETC> {
    if dwdirection as i32 == DATADIR_GET.0 {
      unsafe { SHCreateStdEnumFmtEtc(&self.fmtetc) }
    } else {
      Err(E_NOTIMPL.into())
    }
  }

  fn DAdvise(
    &self,
    _pformatetc: *const FORMATETC,
    _advf: u32,
    _padvsink: Option<&IAdviseSink>,
  ) -> ::windows::core::Result<u32> {
    Err(OLE_E_ADVISENOTSUPPORTED.into())
  }

  fn DUnadvise(&self, _dwconnection: u32) -> ::windows::core::Result<()> {
    Err(OLE_E_ADVISENOTSUPPORTED.into())
  }

  fn EnumDAdvise(&self) -> ::windows::core::Result<IEnumSTATDATA> {
    Err(OLE_E_ADVISENOTSUPPORTED.into())
  }
}

#[allow(non_snake_case)]
impl IDataObjectAsyncCapability_Impl for DragDropObject {
  fn SetAsyncMode(&self, fdoopasync: BOOL) -> ::windows::core::Result<()> {
    self.fdoopasync.replace(fdoopasync.as_bool());
    Ok(())
  }

  fn GetAsyncMode(&self) -> ::windows::core::Result<BOOL> {
    Ok((*self.fdoopasync.borrow()).into())
  }

  fn StartOperation(
    &self,
    _pbcreserved: ::core::option::Option<&IBindCtx>,
  ) -> ::windows::core::Result<()> {
    self.inoperation.replace(true);
    Ok(())
  }

  fn InOperation(&self) -> ::windows::core::Result<BOOL> {
    Ok((*self.inoperation.borrow()).into())
  }

  fn EndOperation(
    &self,
    _hresult: ::windows::core::HRESULT,
    _pbcreserved: ::core::option::Option<&IBindCtx>,
    _dweffects: u32,
  ) -> ::windows::core::Result<()> {
    self.inoperation.replace(false);
    Ok(())
  }
}

fn create_data_object(fmtetc: Vec<FORMATETC>, stgmeds: Vec<STGMEDIUM>) -> DragDropObject {
  DragDropObject {
    fmtetc,
    stgmeds,
    fdoopasync: RefCell::new(false),
    inoperation: RefCell::new(false),
  }
}

#[allow(clippy::single_match)]
fn main() {
  let event_loop = EventLoop::new();

  let mut window = Some(
    WindowBuilder::new()
      .with_title("A fantastic window!")
      .with_inner_size(tao::dpi::LogicalSize::new(300.0, 300.0))
      .with_min_inner_size(tao::dpi::LogicalSize::new(200.0, 200.0))
      .build(&event_loop)
      .unwrap(),
  );

  let drop_source = DragDropClient {};

  event_loop.run(move |event, _, control_flow| {
    *control_flow = ControlFlow::Wait;

    match event {
      Event::WindowEvent {
        event: WindowEvent::CloseRequested,
        window_id: _,
        ..
      } => {
        // drop the window to fire the `Destroyed` event
        window = None;
      }
      Event::WindowEvent {
        event: WindowEvent::Destroyed,
        window_id: _,
        ..
      } => {
        *control_flow = ControlFlow::Exit;
      }
      Event::WindowEvent {
        event:
          WindowEvent::MouseInput {
            state: ElementState::Pressed,
            button: MouseButton::Left,
            ..
          },
        ..
      } => {
        let fmtetc = FORMATETC {
          cfFormat: CF_TEXT.0,
          dwAspect: 0,
          ptd: &DVASPECT_CONTENT as *const _ as *mut _,
          lindex: -1,
          tymed: TYMED_HGLOBAL.0 as _,
        };

        let mut stgmed = STGMEDIUM::default();
        stgmed.tymed = TYMED_HGLOBAL.0 as _;

        let osstr = OsString::from("Hello From TAO DROP");
        let data: Vec<u16> = osstr.encode_wide().chain(Some(0)).collect();
        let handle = global_from_data(data.as_slice()).unwrap();
        stgmed.u.hGlobal = handle;

        let data_object = create_data_object(vec![fmtetc], vec![stgmed]);

        let mut effect = DROPEFFECT_NONE;
        unsafe {
          DoDragDrop(
            Some(&data_object.cast().unwrap()),
            Some(&drop_source.cast().unwrap()),
            DROPEFFECT_COPY | DROPEFFECT_LINK,
            &mut effect,
          )
          .ok()
          .unwrap()
        };
      }
      Event::MainEventsCleared => {
        if let Some(w) = &window {
          w.request_redraw();
        }
      }
      _ => (),
    }
  });
}
