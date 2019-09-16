#![allow(non_camel_case_types, non_snake_case, unused)]

//! Smart & safe rustified WinAPI IUnknown counterpart.
//!

use std::cell::Cell;
use std::convert::{TryFrom, TryInto};
use std::error::Error;

use winapi::shared::guiddef::{IID_NULL, REFIID};
use winapi::shared::minwindef::{LPVOID, PUINT, UINT, WORD};
use winapi::shared::ntdef::{HRESULT, INT, PULONG, ULONG};
use winapi::shared::winerror;
use winapi::shared::wtypes::{BSTR, DATE, VARIANT_BOOL};
use winapi::um::combaseapi::{CoCreateInstance, CoGetClassObject, CLSCTX_ALL};
use winapi::um::oaidl::{
    IDispatch, IDispatchVtbl, DISPID, DISPID_NEWENUM, DISPPARAMS, EXCEPINFO, LPDISPATCH, LPVARIANT,
    SAFEARRAY, VARIANT,
};
use winapi::um::oleauto::{
    SysStringLen, VariantClear, VariantInit, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
    DISPATCH_PROPERTYPUT,
};
use winapi::um::unknwnbase::{IUnknown, IUnknownVtbl, LPUNKNOWN};
use winapi::um::winnt::{LOCALE_USER_DEFAULT, LONG, LPCSTR, LPSTR, WCHAR};
use winapi::{Class, Interface, RIDL};

use crate::auto_com_interface::*;
use crate::smart_variant::*;

pub trait SmartIUnknown {
    fn as_iunknown(&self) -> &IUnknown;
    fn as_iunknown_mut(&mut self) -> &mut IUnknown;

    fn query_interface<T: Interface>(&self) -> Result<AutoCOMInterface<T>, HRESULT> {
        let mut pvoid: LPVOID = std::ptr::null_mut();
        let hresult = unsafe {
            self.as_iunknown()
                .QueryInterface(&<T as winapi::Interface>::uuidof(), &mut pvoid)
        };

        if winerror::SUCCEEDED(hresult) {
            match (pvoid as *mut T).try_into() {
                Ok(x) => Ok(x),
                Err(_) => Err(winerror::E_POINTER),
            }
        } else {
            Err(hresult)
        }
    }

    fn add_ref(&mut self) -> ULONG {
        unsafe { self.as_iunknown_mut().AddRef() }
    }

    fn release(&mut self) -> ULONG {
        unsafe { self.as_iunknown_mut().Release() }
    }
}

impl<T: Interface> SmartIUnknown for T {
    fn as_iunknown(&self) -> &IUnknown {
        unsafe { &*(self as *const Self as *const IUnknown) }
    }

    fn as_iunknown_mut(&mut self) -> &mut IUnknown {
        unsafe { &mut *(self as *mut Self as *mut IUnknown) }
    }
}

impl<T: Interface> SmartIUnknown for AutoCOMInterface<T> {
    fn as_iunknown(&self) -> &IUnknown {
        self.as_iunknown()
    }

    fn as_iunknown_mut(&mut self) -> &mut IUnknown {
        self.as_iunknown_mut()
    }
}

// impl Drop for SmartIUnknown {
//     fn drop(&mut self) {
//         self.release();
//     }
// }

// impl From<LPUNKNOWN> for SmartIUnknown {
//     fn from(x: LPUNKNOWN) -> Self {
//         SmartIUnknown((x).try_into().unwrap())
//     }
// }

// impl<T: Interface> From<AutoCOMInterface<T>> for SmartIUnknown {
//     fn from(x: AutoCOMInterface<T>) -> Self {
//         x.as_iunknown().AddRef();
//         x.as_iunknown_ptr().into()
//     }
// }

#[cfg(test)]
mod tests {
    use super::*;
    use crate::auto_bstr::*;
    use std::convert::TryInto;

    // 1C ComConnector (comcntr.dll) class
    RIDL! {#[uuid(0x181E893D, 0x73A4, 0x4722, 0xB6, 0x1D, 0xD6, 0x04, 0xB3, 0xD6, 0x7D, 0x47)]
    class V8COMConnectorClass;
    }
    pub type LPV8COMCONNECTORCLASS = *mut V8COMConnectorClass;

    RIDL! {#[uuid(0xba4e52bd, 0xdcb2, 0x4bf7, 0xbb, 0x29, 0x84, 0xc1, 0xca, 0x45, 0x6a, 0x8f)]
    interface IV8COMConnector(IV8COMConnectorVtbl): IDispatch(IDispatchVtbl) {
        fn Connect(
            connectString: BSTR,
            conn: *mut LPDISPATCH,
        ) -> HRESULT,
    }}
    pub type LPV8COMCONNECTOR = *mut IV8COMConnector;

    //#[test]
    fn test_AutoCOMInterface_create_instance() {
        let hr = unsafe {
            winapi::um::combaseapi::CoInitializeEx(
                winapi::shared::ntdef::NULL,
                winapi::um::objbase::COINIT_MULTITHREADED,
            )
        };
        assert!(winerror::SUCCEEDED(hr));

        let v8cc = AutoCOMInterface::<IV8COMConnector>::create_instance(
            &<V8COMConnectorClass as Class>::uuidof(),
            std::ptr::null_mut(),
            CLSCTX_ALL,
        )
        .unwrap();

        assert_ne!(v8cc.as_iunknown_ptr(), std::ptr::null_mut());

        let conn1Cdb_bstr: AutoBSTR =
            r#"Srvr="192.168.6.93";Ref="Trade_EP_Today_COPY";"#.try_into().unwrap();

        let mut conn1Cdb: LPDISPATCH = std::ptr::null_mut();

        let hr = unsafe { v8cc.as_inner().Connect(conn1Cdb_bstr.into(), &mut conn1Cdb) };

        assert!(winapi::shared::winerror::SUCCEEDED(hr));

        //        let mut conn1Cdb: AutoCOMInterface<IDispatch> = conn1Cdb.try_into().unwrap();
        let mut conn1Cdb: &mut IUnknown =
            unsafe { &mut *(conn1Cdb as *mut IDispatch as *mut IUnknown) };

        conn1Cdb.add_ref();

        unsafe { winapi::um::combaseapi::CoUninitialize() };
    }
}
