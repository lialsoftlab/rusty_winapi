#![allow(non_camel_case_types, non_snake_case, unused)]

use std::cell::Cell;
use std::convert::{TryFrom, TryInto};

use winapi::shared::ntdef::{NULL, PVOID};
use winapi::shared::wtypes::BSTR;

use super::safe::bstr::*;

/// Container for BSTR-type strings with automatic conversion and memory managment.
pub struct AutoBSTR (Cell<BSTR>);

impl AutoBSTR {
    /// Unconditional freeing allocated memory for BSTR instance now.
    pub fn free(mut self) {
        SysFreeString(self.0.get());
        self.0.set(NULL as BSTR);
    }

    /// Converts ref to AutoBSTR into pointer to BSTR pointer.
    #[inline]
    pub fn as_ptr(&self) -> *const BSTR {
        self.0.as_ptr() as *const BSTR
    }

    /// Converts mutable ref to AutoBSTR into mutable pointer to BSTR pointer.
    #[inline]
    pub fn as_mut_ptr(&mut self) -> *mut BSTR {
        self.0.as_ptr()
    }
}

impl Drop for AutoBSTR {
    fn drop(&mut self) {
        SysFreeString(self.0.get()); // NULL is ok, function just returns.
    }
}

impl TryFrom<&str> for AutoBSTR {
    type Error = super::safe::bstr::SysAllocError;

    /// Try to convert string slice into UTF-16 encoded string, and transform it to new BSTR instance.
    fn try_from(x: &str) -> Result<Self, Self::Error> {
        let utf16_buf: Vec<u16> = x.encode_utf16().collect();
        Ok(AutoBSTR(Cell::new(SysAllocStringLen(&utf16_buf)?)))
    }
}

impl TryFrom<String> for AutoBSTR {
    type Error = super::safe::bstr::SysAllocError;

    /// Try to convert string slice into UTF-16 encoded string, and transform it to new BSTR instance.
    #[inline]
    fn try_from(x: String) -> Result<Self, Self::Error> {
        x.as_str().try_into()
    }
}

impl From<AutoBSTR> for String {
    /// Convert from AutoBSTR instance into UTF-8 encoded Rust String.
    #[inline]
    fn from(x: AutoBSTR) -> Self {
        String::from_utf16_lossy(x.into())
    }
}

impl From<BSTR> for AutoBSTR {
    /// Wrap existing BSTR instance into AutoBSTR with responsibility to free memory on drop.
    #[inline]
    fn from(x: BSTR) -> Self {
        AutoBSTR(Cell::new(x)) 
    }
}

impl From<AutoBSTR> for BSTR {
    /// Convert AutoBSTR instance into BSTR, and mark that we are not resposible to free memory for it anymore.
    fn from(x: AutoBSTR) -> Self {
        let bstr = x.0.get();
        x.0.set(NULL as BSTR);

        bstr
    }
}

impl From<AutoBSTR> for &[u16] {
    /// AutoBSTR instance into [u16] slice reference
    fn from(x: AutoBSTR) -> Self {
        let bstr = x.0.get();
        unsafe { std::slice::from_raw_parts(bstr, SysStringLen(bstr) as usize) }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::convert::TryInto;

    static TEST_LINE: &str = "Test line.\u{0000} Тестовая строка.\u{0000} Testlinie.\u{0000} Ligne de test.\u{0000} Línea de prueba.";

    #[test]
    fn test_AutoBSTR() {
        let auto_bstr: AutoBSTR = TEST_LINE.try_into().unwrap();
        assert_eq!(TEST_LINE, String::from(auto_bstr));

        let test_line_string = String::from(TEST_LINE);
        let mut auto_bstr: AutoBSTR = test_line_string.try_into().unwrap();
        assert_eq!(TEST_LINE, String::from(auto_bstr));

        let test_line_string = String::from(TEST_LINE);
        let mut auto_bstr: AutoBSTR = test_line_string.try_into().unwrap();
        SysFreeString(auto_bstr.0.get());
        unsafe { 
            *auto_bstr.as_mut_ptr() = 0xA5A5A5A5 as BSTR; 
            assert_eq!(*auto_bstr.as_mut_ptr(), *auto_bstr.as_ptr());
        }
        let bstr: BSTR = auto_bstr.into();
        assert_eq!(0xA5A5A5A5 as BSTR, bstr);

    }
}