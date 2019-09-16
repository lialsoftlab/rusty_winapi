#![allow(non_camel_case_types, non_snake_case, unused)]

//! Smart & safe rustified WinAPI IUnknown counterpart.
//!

use std::cell::Cell;
use std::convert::{AsRef, AsMut, TryFrom, TryInto};
use std::error::Error;
use std::ops::{Deref, DerefMut};

use winapi::shared::guiddef::{IID_NULL, REFCLSID, REFIID};
use winapi::shared::minwindef::{DWORD, LPVOID, PUINT, UINT, WORD};
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
use winapi::um::unknwnbase::{IClassFactory, IClassFactoryVtbl, IUnknown, IUnknownVtbl, LPUNKNOWN};
use winapi::um::winnt::{LOCALE_USER_DEFAULT, LONG, LPCSTR, LPSTR, WCHAR};
use winapi::{Class, Interface, RIDL};

use crate::smart_variant::*;

pub struct AutoCOMInterface<T: Interface>(*mut T);

impl<T: Interface> AutoCOMInterface<T> {
    pub fn as_iunknown_ptr(&self) -> LPUNKNOWN {
        unsafe { self.0 as LPUNKNOWN }
    }

    pub fn as_iunknown(&self) -> &IUnknown {
        debug_assert!(
            self.0 != std::ptr::null_mut(),
            "Access to COM interface by uninitialized pointer!"
        );
        unsafe { &*(self.0 as *const IUnknown) }
    }

    pub fn as_iunknown_mut(&mut self) -> &mut IUnknown {
        debug_assert!(
            self.0 != std::ptr::null_mut(),
            "Access to COM interface by uninitialized pointer!"
        );
        unsafe { &mut *(self.0 as *mut IUnknown) }
    }

    pub fn as_inner(&self) -> &T {
        debug_assert!(
            self.0 != std::ptr::null_mut(),
            "Access to COM interface by uninitialized pointer!"
        );
        unsafe { &*self.0 }
    }

    pub fn as_inner_mut(&mut self) -> &mut T {
        debug_assert!(
            self.0 != std::ptr::null_mut(),
            "Access to COM interface by uninitialized pointer!"
        );
        unsafe { &mut *self.0 }
    }

    pub fn unwrap(&mut self) -> *mut T {
        let result = self.0;
        self.0 = std::ptr::null_mut();

        result
    }

    pub fn get_class_object(
        rclsid: REFCLSID,
        dwClsContext: DWORD,
        pvReserved: LPVOID,
    ) -> Result<AutoCOMInterface<T>, HRESULT> {
        let mut pvoid: LPVOID = std::ptr::null_mut();
        let hresult = unsafe {
            CoGetClassObject(
                rclsid,
                dwClsContext,
                pvReserved,
                &<T as winapi::Interface>::uuidof(),
                &mut pvoid,
            )
        };

        if winerror::SUCCEEDED(hresult) {
            Ok(AutoCOMInterface(pvoid as *mut T))
        } else {
            Err(hresult)
        }
    }

    pub fn create_instance(
        rclsid: REFCLSID,
        pUnkOuter: LPUNKNOWN,
        dwClsContext: DWORD,
    ) -> Result<AutoCOMInterface<T>, HRESULT> {
        let mut pvoid: LPVOID = std::ptr::null_mut();
        let hresult = unsafe {
            CoCreateInstance(
                rclsid,
                pUnkOuter,
                dwClsContext,
                &<T as winapi::Interface>::uuidof(),
                &mut pvoid,
            )
        };

        if winerror::SUCCEEDED(hresult) {
            Ok(AutoCOMInterface(pvoid as *mut T))
        } else {
            Err(hresult)
        }
    }
}

impl<T: Interface> Default for AutoCOMInterface<T> {
    fn default() -> Self {
        AutoCOMInterface::<T>(std::ptr::null_mut())
    }
}

impl<T: Interface> Drop for AutoCOMInterface<T> {
    fn drop(&mut self) {
        if self.0 != std::ptr::null_mut() {
            unsafe {
                self.as_iunknown().Release();
            }
        }
    }
}

impl<T: Interface> Deref for AutoCOMInterface<T> {
    type Target = T;

    fn deref(&self) -> &Self::Target {
        unsafe { &*self.0 }
    }
}

impl<T: Interface> DerefMut for AutoCOMInterface<T> {
    fn deref_mut(&mut self) -> &mut Self::Target {
        unsafe { &mut *self.0 }
    }
}

impl<T: Interface> AsRef<T> for AutoCOMInterface<T> {
    fn as_ref(&self) ->&T {
        unsafe{ &*self.0 }
    }
}

impl<T: Interface> AsMut<T> for AutoCOMInterface<T> {
    fn as_mut(&mut self) -> &mut T {
        unsafe { &mut *self.0 }
    }
}

impl<T: Interface> TryFrom<*mut T> for AutoCOMInterface<T> {
    type Error = &'static str;

    fn try_from(x: *mut T) -> Result<Self, Self::Error> {
        if x != std::ptr::null_mut() {
            Ok(AutoCOMInterface(x))
        } else {
            Err("Can't wrap uninitialized COM interface pointer in AutoCOMInterface!")
        }
    }
}

impl TryFrom<SmartVariant> for AutoCOMInterface<IUnknown> {
    type Error = &'static str;

    /// Try to convert string slice into UTF-16 encoded string, and transform it to new BSTR instance.
    #[inline]
    fn try_from(x: SmartVariant) -> Result<Self, Self::Error> {
        match x {
            SmartVariant::IUnknown(x) => unsafe { AutoCOMInterface::try_from(x) },
            _ => Err("SmartVartiant doesn't contains pointer to IUnknown!"),
        }
    }
}

impl TryFrom<SmartVariant> for AutoCOMInterface<IDispatch> {
    type Error = &'static str;

    /// Try to convert string slice into UTF-16 encoded string, and transform it to new BSTR instance.
    #[inline]
    fn try_from(x: SmartVariant) -> Result<Self, Self::Error> {
        match x {
            SmartVariant::IDispatch(x) => unsafe { AutoCOMInterface::try_from(x) },
            _ => Err("SmartVartiant doesn't contains pointer to IDispatch!"),
        }
    }
}

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

    // #[test]
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

        let conn1Cdb: AutoCOMInterface<IDispatch> = conn1Cdb.try_into().unwrap();

        unsafe { winapi::um::combaseapi::CoUninitialize() };
    }
}
