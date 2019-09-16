#![allow(non_camel_case_types, non_snake_case, unused)]

//! Smart & safe rustified WinAPI IClassFactory counterpart.
//!

use std::cell::Cell;
use std::convert::{TryFrom, TryInto};
use std::error::Error;

use winapi::shared::guiddef::{IID_NULL, REFIID};
use winapi::shared::minwindef::{LPVOID, PUINT, UINT, WORD};
use winapi::shared::ntdef::{HRESULT, INT, PULONG, ULONG};
use winapi::shared::winerror;
use winapi::shared::wtypes::{BSTR, DATE, VARIANT_BOOL};
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

use crate::auto_com_interface::*;
use crate::smart_iunknown::*;
use crate::smart_variant::*;

pub trait SmartIClassFactory: SmartIUnknown {
    fn as_iclass_factory(&self) -> &IClassFactory;
    fn as_iclass_factory_mut(&mut self) -> &mut IClassFactory;

    fn create_instance<U: Interface>(
        &self,
        unk_outer: LPUNKNOWN,
    ) -> Result<AutoCOMInterface<U>, HRESULT> {
        let mut pvoid: LPVOID = std::ptr::null_mut();
        let hresult = unsafe {
            self.as_iclass_factory().CreateInstance(
                unk_outer,
                &<U as winapi::Interface>::uuidof(),
                &mut pvoid,
            )
        };

        if winerror::SUCCEEDED(hresult) {
            Ok((pvoid as *mut U).try_into().unwrap())
        } else {
            Err(hresult)
        }
    }

    fn lock_server(&mut self, fLock: bool) -> HRESULT {
        unsafe {
            self.as_iclass_factory_mut()
                .LockServer(if fLock { -1 } else { 0 })
        }
    }
}

impl SmartIClassFactory for IClassFactory {
    fn as_iclass_factory(&self) -> &IClassFactory {
        self
    }

    fn as_iclass_factory_mut(&mut self) -> &mut IClassFactory {
        self
    }
}

impl SmartIClassFactory for AutoCOMInterface<IClassFactory> {
    fn as_iclass_factory(&self) -> &IClassFactory {
        self.as_inner()
    }

    fn as_iclass_factory_mut(&mut self) -> &mut IClassFactory {
        self.as_inner_mut()
    }
}
