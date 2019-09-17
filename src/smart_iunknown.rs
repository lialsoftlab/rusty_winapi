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

    fn add_ref(&self) -> ULONG {
        unsafe { self.as_iunknown().AddRef() }
    }

    fn release(&self) -> ULONG {
        unsafe { self.as_iunknown().Release() }
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

#[cfg(test)]
mod tests {}
