#![allow(non_camel_case_types, non_snake_case, unused)]

//! Smart & safe rustified WinAPI VARIANT counterpart.
//!
//! See also:
//!
//! * https://docs.microsoft.com/en-us/windows/win32/winauto/variant-structure
//! * https://docs.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-variant
//! * https://docs.microsoft.com/ru-ru/previous-versions/windows/desktop/automat/variant-manipulation-functions
//! * https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/3fe7db9f-5803-4dc4-9d14-5425d3f5461f
//! * https://docs.microsoft.com/en-us/windows/win32/api/wtypes/ne-wtypes-varenum
//!

use std::any::Any;
use std::cell::Cell;
use std::convert::{AsMut, AsRef, TryFrom};

use winapi::shared::minwindef::UINT;
use winapi::shared::ntdef::*;
use winapi::shared::wtypes::*;
use winapi::shared::wtypesbase::*;
use winapi::um::oaidl::*;
use winapi::um::unknwnbase::*;

use crate::auto_bstr::AutoBSTR;

#[derive(Clone, Debug, PartialEq)]
pub enum SmartVariant {
    Empty,
    Int2(i16),
    Int4(i32),
    Real4(f32),
    Real8(f64),
    //Currency(CY),
    Date(f64),
    Text(String),
    IDispatch(LPDISPATCH),
    ErrorCode(i32), // SCODE
    Bool(bool),
    Variant(LPVARIANT),
    IUnknown(LPUNKNOWN),
    //Decimal(i128),
    Int1(i8),
    UInt1(u8),
    UInt2(u16),
    UInt4(u32),
    Int(i32),
    UInt(u32),
    //Record(LPRECORD),
    Array(LPSAFEARRAY),
    ByRef(PVOID), // mask value?
}

pub struct AutoVariant(Cell<VARIANT>);

impl AutoVariant {
    #[inline]
    pub fn new() -> AutoVariant {
        AutoVariant(Cell::new(VARIANT::default())) // New zeroed with vt == VT_EMPTY
    }

    #[inline]
    pub fn clear(&mut self) -> HRESULT {
        unsafe {
            if self.vtype() != VT_EMPTY {
                let hresult = winapi::um::oleauto::VariantClear(self.0.get_mut());
                *self.vtype_mut() = VT_EMPTY as u16;

                hresult
            } else {
                0
            }
        }
    }

    #[inline]
    pub fn vtype(&self) -> VARENUM {
        unsafe { self.0.get().n1.n2().vt as VARENUM }
    }

    #[inline]
    pub fn vtype_mut(&mut self) -> &mut u16 {
        unsafe { &mut self.0.get_mut().n1.n2_mut().vt }
    }

    #[inline]
    pub fn data(&self) -> &VARIANT_n3 {
        unsafe { &(*self.0.as_ptr()).n1.n2().n3 }
    }

    #[inline]
    pub fn data_mut(&mut self) -> &mut VARIANT_n3 {
        unsafe { &mut self.0.get_mut().n1.n2_mut().n3 }
    }

    pub fn value(&self) -> &dyn Any {
        unsafe {
            match self.vtype() {
                VT_I2 => self.data().iVal(),           // A 2-byte integer.
                VT_I4 => self.data().lVal(),           // A 4-byte integer.
                VT_R4 => self.data().fltVal(),         // A 4-byte real.
                VT_R8 => self.data().dblVal(),         // An 8-byte real.
                VT_CY => self.data().cyVal(),          // Currency. (i64)
                VT_DATE => self.data().date(),         // A date. (f64)
                VT_BSTR => self.data().bstrVal(),      // A string.
                VT_DISPATCH => self.data().pdispVal(), //An IDispatch pointer.
                VT_ERROR => self.data().scode(),       // An SCODE value. (i32)
                VT_BOOL => self.data().boolVal(), //A Boolean value. True is -1 and false is 0. (i16)
                VT_VARIANT => self.data().pvarVal(), // A variant pointer.
                VT_UNKNOWN => self.data().punkVal(), // An IUnknown pointer.
                VT_DECIMAL => self.data().pdecVal(), // A 16-byte fixed-pointer value.
                VT_I1 => self.data().cVal(),      // A character. (i8)
                VT_UI1 => self.data().bVal(),     // An unsigned character. (u8)
                VT_UI2 => self.data().uiVal(),    // An unsigned short. (u16)
                VT_UI4 => self.data().ulVal(),    // An unsigned long.  (u32)
                VT_INT => self.data().intVal(),   // An integer. (i32)
                VT_UINT => self.data().uintVal(), // An unsigned integer. (u32)
                VT_RECORD => self.data().n4(),    // A user-defined type.
                VT_ARRAY => self.data().parray(), // A SAFEARRAY pointer.
                VT_BYREF => self.data().byref(),  // A void pointer for local use.
                _ => self.data(),
            }
        }
    }

    pub fn value_mut(&mut self) -> &mut dyn Any {
        unsafe {
            match self.vtype() {
                VT_I2 => self.data_mut().iVal_mut(),      // A 2-byte integer.
                VT_I4 => self.data_mut().lVal_mut(),      // A 4-byte integer.
                VT_R4 => self.data_mut().fltVal_mut(),    // A 4-byte real.
                VT_R8 => self.data_mut().dblVal_mut(),    // An 8-byte real.
                VT_CY => self.data_mut().cyVal_mut(),     // Currency. (i64)
                VT_DATE => self.data_mut().date_mut(),    // A date. (f64)
                VT_BSTR => self.data_mut().bstrVal_mut(), // A string.
                VT_DISPATCH => self.data_mut().pdispVal_mut(), //An IDispatch pointer.
                VT_ERROR => self.data_mut().scode_mut(),  // An SCODE value. (i32)
                VT_BOOL => self.data_mut().boolVal_mut(), //A Boolean value. True is -1 and false is 0. (i16)
                VT_VARIANT => self.data_mut().pvarVal_mut(), // A variant pointer.
                VT_UNKNOWN => self.data_mut().punkVal_mut(), // An IUnknown pointer.
                VT_DECIMAL => self.data_mut().pdecVal_mut(), // A 16-byte fixed-pointer value.
                VT_I1 => self.data_mut().cVal_mut(),      // A character. (i8)
                VT_UI1 => self.data_mut().bVal_mut(),     // An unsigned character. (u8)
                VT_UI2 => self.data_mut().uiVal_mut(),    // An unsigned short. (u16)
                VT_UI4 => self.data_mut().ulVal_mut(),    // An unsigned long.  (u32)
                VT_INT => self.data_mut().intVal_mut(),   // An integer. (i32)
                VT_UINT => self.data_mut().uintVal_mut(), // An unsigned integer. (u32)
                VT_RECORD => self.data_mut().n4_mut(),    // A user-defined type.
                VT_ARRAY => self.data_mut().parray_mut(), // A SAFEARRAY pointer.
                VT_BYREF => self.data_mut().byref_mut(),  // A void pointer for local use.
                _ => self.data_mut(),
            }
        }
    }

    pub fn value_set<T: Any>(mut self, value: &T) -> Self {
        let value = value as &dyn Any;

        self.clear();

        if let Some(&n_i16) = value.downcast_ref::<i16>() {
            unsafe {
                *self.vtype_mut() = VT_I2 as u16;
                *self.data_mut().iVal_mut() = n_i16;
            }
        } else if let Some(&n_i32) = value.downcast_ref::<i32>() {
            unsafe {
                *self.vtype_mut() = VT_I4 as u16;
                *self.data_mut().lVal_mut() = n_i32;
            }
        } else if let Some(&n_f32) = value.downcast_ref::<f32>() {
            unsafe {
                *self.vtype_mut() = VT_R4 as u16;
                *self.data_mut().fltVal_mut() = n_f32;
            }
        } else if let Some(&n_f64) = value.downcast_ref::<f64>() {
            unsafe {
                *self.vtype_mut() = VT_R8 as u16;
                *self.data_mut().dblVal_mut() = n_f64;
            }
        } else if let Some(&cy) = value.downcast_ref::<CY>() {
            unsafe {
                *self.vtype_mut() = VT_CY as u16;
                *self.data_mut().cyVal_mut() = cy;
            }
        } else if let Some(&date) = value.downcast_ref::<DATE>() {
            unsafe {
                *self.vtype_mut() = VT_DATE as u16;
                *self.data_mut().date_mut() = date;
            }
        } else if let Some(&bstr) = value.downcast_ref::<BSTR>() {
            unsafe {
                *self.vtype_mut() = VT_BSTR as u16;
                *self.data_mut().bstrVal_mut() = bstr;
            }
        } else if let Some(&pdisp) = value.downcast_ref::<LPDISPATCH>() {
            unsafe {
                *self.vtype_mut() = VT_DISPATCH as u16;
                *self.data_mut().pdispVal_mut() = pdisp;
            }
        } else if let Some(&error) = value.downcast_ref::<SCODE>() {
            unsafe {
                *self.vtype_mut() = VT_ERROR as u16;
                *self.data_mut().scode_mut() = error;
            }
        } else if let Some(&boolean) = value.downcast_ref::<bool>() {
            unsafe {
                *self.vtype_mut() = VT_BOOL as u16;
                *self.data_mut().boolVal_mut() = if boolean { VARIANT_TRUE } else { VARIANT_FALSE };
            }
        } else if let Some(&pvar) = value.downcast_ref::<LPVARIANT>() {
            unsafe {
                *self.vtype_mut() = VT_VARIANT as u16;
                *self.data_mut().pvarVal_mut() = pvar;
            }
        } else if let Some(&punk) = value.downcast_ref::<LPUNKNOWN>() {
            unsafe {
                *self.vtype_mut() = VT_UNKNOWN as u16;
                *self.data_mut().punkVal_mut() = punk;
            }
        } else if let Some(&pdec) = value.downcast_ref::<LPDECIMAL>() {
            unsafe {
                *self.vtype_mut() = VT_DECIMAL as u16;
                *self.data_mut().pdecVal_mut() = pdec;
            }
        } else if let Some(&n_i8) = value.downcast_ref::<i8>() {
            unsafe {
                *self.vtype_mut() = VT_I1 as u16;
                *self.data_mut().cVal_mut() = n_i8;
            }
        } else if let Some(&n_u8) = value.downcast_ref::<u8>() {
            unsafe {
                *self.vtype_mut() = VT_UI1 as u16;
                *self.data_mut().bVal_mut() = n_u8;
            }
        } else if let Some(&n_u16) = value.downcast_ref::<u16>() {
            unsafe {
                *self.vtype_mut() = VT_UI2 as u16;
                *self.data_mut().uiVal_mut() = n_u16;
            }
        } else if let Some(&n_u32) = value.downcast_ref::<u32>() {
            unsafe {
                *self.vtype_mut() = VT_UI4 as u16;
                *self.data_mut().ulVal_mut() = n_u32;
            }
        } else if let Some(&n_i32) = value.downcast_ref::<INT>() {
            unsafe {
                *self.vtype_mut() = VT_INT as u16;
                *self.data_mut().intVal_mut() = n_i32;
            }
        } else if let Some(&n_u32) = value.downcast_ref::<UINT>() {
            unsafe {
                *self.vtype_mut() = VT_UINT as u16;
                *self.data_mut().uintVal_mut() = n_u32;
            }
        } else if let Some(&rec) = value.downcast_ref::<__tagBRECORD>() {
            unsafe {
                *self.vtype_mut() = VT_RECORD as u16;
                *self.data_mut().n4_mut() = rec;
            }
        } else if let Some(&parr) = value.downcast_ref::<LPSAFEARRAY>() {
            unsafe {
                *self.vtype_mut() = VT_ARRAY as u16;
                *self.data_mut().parray_mut() = parr;
            }
        } else if let Some(&pvoid) = value.downcast_ref::<PVOID>() {
            unsafe {
                *self.vtype_mut() = VT_BYREF as u16;
                *self.data_mut().byref_mut() = pvoid;
            }
        } else {
            unsafe {
                *self.vtype_mut() = VT_EMPTY as u16;
                *self.data_mut().llVal_mut() = 0;
            }
        }

        self
    }
}

impl Drop for AutoVariant {
    #[inline]
    fn drop(&mut self) {
        self.clear();
    }
}

impl From<AutoVariant> for VARIANT {
    #[inline]
    fn from(x: AutoVariant) -> Self {
        let result = x.0.get();
        unsafe { (*x.0.as_ptr()).n1.n2_mut().vt = VT_EMPTY as u16 };

        result
    }
}

impl From<VARIANT> for AutoVariant {
    #[inline]
    fn from(x: VARIANT) -> Self {
        AutoVariant(Cell::new(x))
    }
}

impl From<AutoVariant> for SmartVariant {
    #[inline]
    fn from(x: AutoVariant) -> Self {
        let vtype = x.vtype();

        unsafe {
            (*x.0.as_ptr()).n1.n2_mut().vt = VT_EMPTY as u16;
            match vtype {
                VT_EMPTY => SmartVariant::Empty,
                VT_I2 => SmartVariant::Int2(*x.data().iVal()), // A 2-byte integer.
                VT_I4 => SmartVariant::Int4(*x.data().lVal()), // A 4-byte integer.
                VT_R4 => SmartVariant::Real4(*x.data().fltVal()), // A 4-byte real.
                VT_R8 => SmartVariant::Real8(*x.data().dblVal()), // An 8-byte real.
                //VT_CY => SmartVariant::Currency(*x.data().cyVal()), // Currency. (i64)
                VT_DATE => SmartVariant::Date(*x.data().date()), // A date. (f64)
                VT_BSTR => SmartVariant::Text(AutoBSTR::from(*x.data().bstrVal()).into()), // A string.
                VT_DISPATCH => SmartVariant::IDispatch(*x.data().pdispVal()), //An IDispatch pointer.
                VT_ERROR => SmartVariant::ErrorCode(*x.data().scode()), // An SCODE value. (i32)
                VT_BOOL => SmartVariant::Bool(*x.data().boolVal() == -1), //A Boolean value. True is -1 and false is 0. (i16)
                VT_VARIANT => SmartVariant::Variant(*x.data().pvarVal()), // A variant pointer.
                VT_UNKNOWN => SmartVariant::IUnknown(*x.data().punkVal()), // An IUnknown pointer.
                //VT_DECIMAL => SmartVariant::Decimal(*x.data().pdecVal()), // A 16-byte fixed-pointer value.
                VT_I1 => SmartVariant::Int1(*x.data().cVal()), // A character. (i8)
                VT_UI1 => SmartVariant::UInt1(*x.data().bVal()), // An unsigned character. (u8)
                VT_UI2 => SmartVariant::UInt2(*x.data().uiVal()), // An unsigned short. (u16)
                VT_UI4 => SmartVariant::UInt4(*x.data().ulVal()), // An unsigned long.  (u32)
                VT_INT => SmartVariant::Int(*x.data().intVal()), // An integer. (i32)
                VT_UINT => SmartVariant::UInt(*x.data().uintVal()), // An unsigned integer. (u32)
                //VT_RECORD => SmartVariant::Record(*x.data().n4()), // A user-defined type.
                VT_ARRAY => SmartVariant::Array(*x.data().parray()), // A SAFEARRAY pointer.
                VT_BYREF => SmartVariant::ByRef(*x.data().byref()), // A void pointer for local use.
                _ => panic!("Unsupported type for VARIANT"),
            }
        }
    }
}

impl From<VARIANT> for SmartVariant {
    #[inline]
    fn from(x: VARIANT) -> Self {
        AutoVariant::from(x).into()
    }
}

impl From<SmartVariant> for AutoVariant {
    #[inline]
    fn from(x: SmartVariant) -> Self {
        let mut result = AutoVariant::new();
        unsafe {
            match x {
                SmartVariant::Empty => result,
                SmartVariant::Int2(x) => {
                    *result.vtype_mut() = VT_I2 as u16;
                    *result.data_mut().iVal_mut() = x;
                    result
                } // A 2-byte integer.
                SmartVariant::Int4(x) => {
                    *result.vtype_mut() = VT_I4 as u16;
                    *result.data_mut().lVal_mut() = x;
                    result
                } // A 4-byte integer.
                SmartVariant::Real4(x) => {
                    *result.vtype_mut() = VT_R4 as u16;
                    *result.data_mut().fltVal_mut() = x;
                    result
                } // A 4-byte real.
                SmartVariant::Real8(x) => {
                    *result.vtype_mut() = VT_R8 as u16;
                    *result.data_mut().dblVal_mut() = x;
                    result
                } // An 8-byte real.
                //SmartVariant::Currency(x) => { *result.vtype_mut() = VT_CY as u16; *result.data_mut().cyVal_mut() = x as CY }, // Currency. (i64)
                SmartVariant::Date(x) => {
                    *result.vtype_mut() = VT_DATE as u16;
                    *result.data_mut().date_mut() = x;
                    result
                } // A date. (f64)
                SmartVariant::Text(x) => {
                    *result.vtype_mut() = VT_BSTR as u16;
                    *result.data_mut().bstrVal_mut() = AutoBSTR::try_from(x).unwrap().into();
                    result
                } // A string.
                SmartVariant::IDispatch(x) => {
                    *result.vtype_mut() = VT_DISPATCH as u16;
                    *result.data_mut().pdispVal_mut() = x;
                    result
                } //An IDispatch pointer.
                SmartVariant::ErrorCode(x) => {
                    *result.vtype_mut() = VT_ERROR as u16;
                    *result.data_mut().scode_mut() = x;
                    result
                } // An SCODE value. (i32)
                SmartVariant::Bool(x) => {
                    *result.vtype_mut() = VT_BOOL as u16;
                    *result.data_mut().boolVal_mut() = if x { -1 } else { 0 };
                    result
                } //A Boolean value. True is -1 and false is 0. (i16)
                SmartVariant::Variant(x) => {
                    *result.vtype_mut() = VT_VARIANT as u16;
                    *result.data_mut().pvarVal_mut() = x;
                    result
                } // A variant pointer.
                SmartVariant::IUnknown(x) => {
                    *result.vtype_mut() = VT_UNKNOWN as u16;
                    *result.data_mut().punkVal_mut() = x;
                    result
                } // An IUnknown pointer.
                //SmartVariant::Decimal(x) => { *result.vtype_mut() = VT_DECIMAL as u16; *result.data_mut().pdecVal_mut() = x; result }, // A 16-byte fixed-pointer value.
                SmartVariant::Int1(x) => {
                    *result.vtype_mut() = VT_I1 as u16;
                    *result.data_mut().cVal_mut() = x;
                    result
                } // A character. (i8)
                SmartVariant::UInt1(x) => {
                    *result.vtype_mut() = VT_UI1 as u16;
                    *result.data_mut().bVal_mut() = x;
                    result
                } // An unsigned character. (u8)
                SmartVariant::UInt2(x) => {
                    *result.vtype_mut() = VT_UI2 as u16;
                    *result.data_mut().uiVal_mut() = x;
                    result
                } // An unsigned short. (u16)
                SmartVariant::UInt4(x) => {
                    *result.vtype_mut() = VT_UI4 as u16;
                    *result.data_mut().ulVal_mut() = x;
                    result
                } // An unsigned long.  (u32)
                SmartVariant::Int(x) => {
                    *result.vtype_mut() = VT_INT as u16;
                    *result.data_mut().intVal_mut() = x;
                    result
                } // An integer. (i32)
                SmartVariant::UInt(x) => {
                    *result.vtype_mut() = VT_UINT as u16;
                    *result.data_mut().uintVal_mut() = x;
                    result
                } // An unsigned integer. (u32)
                //SmartVariant::Record(x) => { *result.vtype_mut() = VT_RECORD as u16; *result.data_mut().n4_mut() = x; result }, // A user-defined type.
                SmartVariant::Array(x) => {
                    *result.vtype_mut() = VT_ARRAY as u16;
                    *result.data_mut().parray_mut() = x;
                    result
                } // A SAFEARRAY pointer.
                SmartVariant::ByRef(x) => {
                    *result.vtype_mut() = VT_BYREF as u16;
                    *result.data_mut().byref_mut() = x;
                    result
                } // A void pointer for local use.
            }
        }
    }
}

impl From<SmartVariant> for VARIANT {
    #[inline]
    fn from(x: SmartVariant) -> Self {
        AutoVariant::from(x).into()
    }
}

#[cfg(test)]
mod tests {
    use std::convert::{TryFrom, TryInto};
    use winapi::shared::minwindef::UINT;
    use winapi::shared::ntdef::*;
    use winapi::shared::wtypes::*;
    use winapi::shared::wtypesbase::*;
    use winapi::um::oaidl::*;
    use winapi::um::unknwnbase::*;

    use super::*;

    #[test]
    fn test1() {}
}
