#![allow(non_camel_case_types, non_snake_case, unused)]

//! Smart & safe rustified WinAPI IUnknown counterpart.
//!

use std::cell::Cell;
use std::clone::Clone;
use std::cmp::PartialEq;
use std::convert::{AsMut, AsRef, TryFrom, TryInto};
use std::error::Error;
use std::fmt::{self, Debug};
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

use crate::smart_iunknown::SmartIUnknown;
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

    pub fn unwrap(&self) -> *mut T {
        if self.0 != std::ptr::null_mut() {
            self.add_ref();
        }

        self.0
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

impl<T: Interface> Clone for AutoCOMInterface<T> {
    fn clone(&self) -> Self {
        if self.0 != std::ptr::null_mut() {
            unsafe {
                self.add_ref();
            }
        }

        AutoCOMInterface::<T>(self.0)
    }
}

impl<T: Interface> Debug for AutoCOMInterface<T> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "AutoCOMInterface(0x{:X})", self.0 as usize)
    }
}

impl<T: Interface> Drop for AutoCOMInterface<T> {
    fn drop(&mut self) {
        if self.0 != std::ptr::null_mut() {
            unsafe {
                self.release();
            }
        }
    }
}

impl<T: Interface> PartialEq for AutoCOMInterface<T> {
    fn eq(&self, other: &AutoCOMInterface<T>) -> bool {
        self.0 == other.0
    }
}

impl<T: Interface> PartialEq<*mut T> for AutoCOMInterface<T> {
    fn eq(&self, other: &*mut T) -> bool {
        self.0 == *other
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
    fn as_ref(&self) -> &T {
        unsafe { &*self.0 }
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

#[cfg(test)]
mod tests {}
