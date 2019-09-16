#![allow(non_camel_case_types, non_snake_case)]

//! Safe counterparts of WinAPI functions for BSTR strings management.
//!
//! Take a look at [`AutoBSTR`] instead of direct use of this functions, for automatic handling and conversion from/to [`String`].
//!
//! See also: [BSTR] at MSDN, [Eric’s Complete Guide To BSTR Semantics], and [BSTR specification].
//!
//! [`AutoBSTR`]: ../../auto_bstr/struct.AutoBSTR.html
//! [Eric’s Complete Guide To BSTR Semantics]: https://blogs.msdn.microsoft.com/ericlippert/2003/09/12/erics-complete-guide-to-bstr-semantics/
//! [BSTR]: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/automat/bstr/
//! [BSTR specification]: https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-dtyp/692a42a9-06ce-4394-b9bc-5d2a50440168
//! [`String`]: https://doc.rust-lang.org/std/string/struct.String.html
//!

use std::convert::TryFrom;

use winapi::shared::minwindef::{BOOL, TRUE, UINT};
use winapi::shared::ntdef::{NULL, PVOID};
use winapi::shared::wtypes::BSTR;

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum SysAllocError {
    BStrAllocationError,
    InvalidPointerError,
    NullTerminatedStringRequiredError,
    SourceStringTooLongError,
}

/// Allocates a new [BSTR] string and copies the passed UTF-16 null-terminated source string into it.
///
/// If source is a zero-length string, returns a new zero-length [BSTR] string.
///
/// See also [MSDN SysAllocString] description.
///
/// # Errors
///
/// * If source is not null-terminated, returns [`NullTerminatedStringRequiredError`].
/// * If insufficient memory exists, returns [`BStrAllocationError`].
///
/// # Examples
///
/// ```
///     use rusty_winapi::safe::bstr::{SysAllocString, SysFreeString, SysStringLen};
///
///     let test_string: Vec<u16> = "Test string.\u{0000} (buffer may be longer)".encode_utf16().collect();
///     let bstr = SysAllocString(&test_string).expect("BSTR");
///     let bstr_slice = unsafe { std::slice::from_raw_parts(bstr, SysStringLen(bstr) as usize) };
///
///     assert_eq!("Test string.", String::from_utf16_lossy(bstr_slice));
///     SysFreeString(bstr);
/// ```
///
/// [BSTR]: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/automat/bstr/
/// [`BStrAllocationError`]: enum.SysAllocError.html#variant.BStrAllocationError
/// [MSDN SysAllocString]: https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysallocstring
/// [`NullTerminatedStringRequiredError`]: enum.SysAllocError.html#variant.NullTerminatedStringRequiredError
pub fn SysAllocString(src: &[u16]) -> Result<BSTR, SysAllocError> {
    // Source string must be null-terminated.
    if !src.iter().any(|x| x == &0) {
        return Err(SysAllocError::NullTerminatedStringRequiredError);
    };

    unsafe {
        match winapi::um::oleauto::SysAllocString(src.as_ptr()) as PVOID {
            NULL => Err(SysAllocError::BStrAllocationError),
            x => Ok(x as BSTR),
        }
    }
}

/// Reallocates a previously allocated string to be the size of a UTF-16 null-terminated source string and copies the source
/// string into the reallocated memory. Then frees the old [BSTR].
///
/// If source is a zero-length string, returns a zero-length [BSTR].
///
/// See also [MSDN SysReAllocString] description.
///
/// # Errors
///
/// * If bstr is NULL or points into source memory range, returns [`InvalidPointerError`].
/// * If source is not null-terminated, returns [`NullTerminatedStringRequiredError`].
/// * If insufficient memory exists, returns [`BStrAllocationError`].
///
/// # Examples
///
/// ```
/// use rusty_winapi::safe::bstr::{SysAllocString, SysFreeString, SysReAllocString, SysStringLen};
///
/// let test_string: Vec<u16> = "Test string.\u{0000} (buffer may be longer)".encode_utf16().collect();
/// let bstr = SysAllocString(&test_string).expect("BSTR");
/// let test_string: Vec<u16> = "New test string.\u{0000}".encode_utf16().collect();
/// let bstr = SysReAllocString(bstr, &test_string).expect("BSTR");
/// let bstr_slice = unsafe { std::slice::from_raw_parts(bstr, SysStringLen(bstr) as usize) };
///
/// assert_eq!("New test string.", String::from_utf16_lossy(bstr_slice));
/// SysFreeString(bstr);
/// ```
///
/// [BSTR]: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/automat/bstr/
/// [`BStrAllocationError`]: enum.SysAllocError.html#variant.BStrAllocationError
/// [`InvalidPointerError`]: enum.SysAllocError.html#variant.InvalidPointerError
/// [MSDN SysReAllocString]: https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysreallocstring
/// [`NullTerminatedStringRequiredError`]: enum.SysAllocError.html#variant.NullTerminatedStringRequiredError
pub fn SysReAllocString(bstr: BSTR, src: &[u16]) -> Result<BSTR, SysAllocError> {
    // If pbstr is NULL, there will be an access violation and the program will crash.
    if bstr as PVOID == NULL {
        return Err(SysAllocError::InvalidPointerError);
    };

    // Source string must be null-terminated.
    if !src.iter().any(|x| x == &0) {
        return Err(SysAllocError::NullTerminatedStringRequiredError);
    };

    // The address passed in bstr cannot be part of the string passed in src, or unexpected results may occur.
    if bstr_src_intersection(bstr, src) {
        return Err(SysAllocError::InvalidPointerError);
    };

    let mut result = bstr;
    unsafe {
        match winapi::um::oleauto::SysReAllocString(&mut result, src.as_ptr()) as BOOL {
            TRUE => Ok(result),
            _ => Err(SysAllocError::BStrAllocationError),
        }
    }
}

/// Allocates a new [BSTR] string, copies the passed UTF-16 source string into it (max up to std::u32::MAX characters),
/// and appends a null-terminating character.
///
/// The string can contain embedded null characters and does not need to end with a NULL.
///
/// See also [MSDN SysAllocStringLen] description.
///
/// # Errors
///
/// * If there is insufficient memory to complete the operation, returns [`BStrAllocationError`].
/// * If source string length is more than std::u32::MAX, returns [`SourceStringTooLongError`].
///
/// # Examples
///
/// ```
/// use rusty_winapi::safe::bstr::{SysAllocStringLen, SysFreeString, SysStringLen};
///
/// let test_string: Vec<u16> = "Test string.".encode_utf16().collect();
/// let bstr = SysAllocStringLen(&test_string).expect("BSTR");
/// let bstr_slice = unsafe { std::slice::from_raw_parts(bstr, SysStringLen(bstr) as usize) };
///
/// assert_eq!("Test string.", String::from_utf16_lossy(bstr_slice));
/// SysFreeString(bstr);
/// ```
///
/// [BSTR]: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/automat/bstr/
/// [`BStrAllocationError`]: enum.SysAllocError.html#variant.BStrAllocationError
/// [MSDN SysAllocStringLen]: https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysallocstringlen
/// [`SourceStringTooLongError`]: enum.SysAllocError.html#variant.SourceStringTooLongError
pub fn SysAllocStringLen(src: &[u16]) -> Result<BSTR, SysAllocError> {
    let len: u32 = match TryFrom::try_from(src.len()) {
        Ok(x) => x,
        Err(_) => return Err(SysAllocError::SourceStringTooLongError),
    };

    unsafe {
        match winapi::um::oleauto::SysAllocStringLen(src.as_ptr(), len) as PVOID {
            NULL => Err(SysAllocError::NullTerminatedStringRequiredError),
            x => Ok(x as BSTR),
        }
    }
}

/// Reallocates a previously allocated [BSTR] string to be the size of a UTF-16 source string and copies the source
/// string into the reallocated memory (max up to std::u32::MAX characters). Then frees the old BSTR.
///
/// The string can contain embedded null characters and does not need to end with a NULL.
/// If source is a zero-length string, returns a zero-length [BSTR].
///
/// See also [MSDN SysReAllocStringLen] description.
///
/// # Errors
///
/// * If bstr is NULL or points into source memory range, returns [`InvalidPointerError`].
/// * If source string length is more than std::u32::MAX, returns [`SourceStringTooLongError`].
/// * If insufficient memory exists, returns [`BStrAllocationError`].
///
/// # Examples
///
/// ```
/// use rusty_winapi::safe::bstr::{SysAllocStringLen, SysFreeString, SysReAllocStringLen, SysStringLen};
///
/// let test_string: Vec<u16> = "Test string.".encode_utf16().collect();
/// let bstr = SysAllocStringLen(&test_string).expect("BSTR");
/// let test_string: Vec<u16> = "New test string.".encode_utf16().collect();
/// let bstr = SysReAllocStringLen(bstr, &test_string).expect("BSTR");
/// let bstr_slice = unsafe { std::slice::from_raw_parts(bstr, SysStringLen(bstr) as usize) };
///
/// assert_eq!("New test string.", String::from_utf16_lossy(bstr_slice));
/// SysFreeString(bstr);
/// ```
///
/// [BSTR]: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/automat/bstr/
/// [`BStrAllocationError`]: enum.SysAllocError.html#variant.BStrAllocationError
/// [`InvalidPointerError`]: enum.SysAllocError.html#variant.InvalidPointerError
/// [MSDN SysReAllocStringLen]: https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysreallocstringlen
/// [`SourceStringTooLongError`]: enum.SysAllocError.html#variant.SourceStringTooLongError
pub fn SysReAllocStringLen(bstr: BSTR, src: &[u16]) -> Result<BSTR, SysAllocError> {
    let len: u32 = match TryFrom::try_from(src.len()) {
        Ok(x) => x,
        Err(_) => return Err(SysAllocError::SourceStringTooLongError),
    };

    // If pbstr is NULL, there will be an access violation and the program will crash.
    if bstr as PVOID == NULL {
        return Err(SysAllocError::InvalidPointerError);
    };

    // The address passed in bstr cannot be part of the string passed in src, or unexpected results may occur.
    if bstr_src_intersection(bstr, src) {
        return Err(SysAllocError::InvalidPointerError);
    };

    let mut result = bstr;
    unsafe {
        match winapi::um::oleauto::SysReAllocStringLen(&mut result, src.as_ptr(), len) as BOOL {
            TRUE => Ok(result),
            _ => Err(SysAllocError::BStrAllocationError),
        }
    }
}

/// Returns the length of a [BSTR].
///
/// The number of characters in bstr, not including the terminating NULL character. If bstr is NULL the return value is zero.
/// The returned value may be different from strlen(bstr) if the [BSTR] contains embedded NULL characters.
/// This function always returns the number of characters specified in the cch parameter of the [MSDN SysAllocStringLen]
/// function used to allocate the [BSTR].
///
/// See also [MSDN SysStringLen] description.
///
/// # Examples
///
/// ```
/// use rusty_winapi::safe::bstr::{SysAllocStringLen, SysFreeString, SysStringLen};
///
/// let test_string: Vec<u16> = "Test string.".encode_utf16().collect();
/// let bstr = SysAllocStringLen(&test_string).expect("BSTR");
///
/// assert_eq!(12, SysStringLen(bstr));
/// SysFreeString(bstr);
/// ```
///
/// [BSTR]: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/automat/bstr/
/// [MSDN SysAllocStringLen]: https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysallocstringlen
/// [MSDN SysStringLen]: https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysstringlen
#[inline]
pub fn SysStringLen(bstr: BSTR) -> UINT {
    unsafe { winapi::um::oleauto::SysStringLen(bstr) }
}

/// Deallocates a [BSTR] string allocated previously by [`SysAllocString`], [`SysAllocStringByteLen`], [`SysReAllocString`],
/// [`SysAllocStringLen`], or [`SysReAllocStringLen`].
///
/// This function does not return a value. If this parameter is NULL, the function simply returns.
///
/// See also [MSDN SysFreeString] description.
///
/// # Examples
///
/// ```
/// use rusty_winapi::safe::bstr::{SysAllocStringLen, SysFreeString, SysStringLen};
///
/// let test_string: Vec<u16> = "Test string.".encode_utf16().collect();
/// let bstr = SysAllocStringLen(&test_string).expect("BSTR");
/// SysFreeString(bstr);
/// ```
///
/// [BSTR]: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/automat/bstr/
/// [MSDN SysFreeString]: https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysfreestring
/// [`SysAllocString`]: fn.SysAllocString.html
/// [`SysAllocStringByteLen`]: fn.SysAllocStringByteLen.html
/// [`SysReAllocString`]: fn.SysReAllocString.html
/// [`SysAllocStringLen`]: fn.SysAllocStringLen.html
/// [`SysReAllocStringLen`]: fn.SysReAllocStringLen.html
#[inline]
pub fn SysFreeString(bstr: BSTR) {
    unsafe {
        winapi::um::oleauto::SysFreeString(bstr);
    }
}

#[inline(always)]
fn bstr_src_intersection(bstr: BSTR, src: &[u16]) -> bool {
    const SIZE_OF_U16: isize = std::mem::size_of::<u16>() as isize;

    if src.len() == 0 || src.len() > std::u32::MAX as usize || bstr as PVOID == NULL {
        return false;
    }; // src.len() == 0 is equal to NULL pointer by meaning.

    // Real BSTR buffer len = {4-byte LenCounter} + SysStringLen() * 2 + 2-byte EOL 0x0000 marker.
    let bstr_start_ptr = bstr as *const u8;
    let bstr_start_ptr = unsafe { bstr_start_ptr.offset(-4) }; // taking in account 32-bit buffer size counter before buffer.
    let bstr_end_ptr = unsafe {
        bstr_start_ptr.offset(4 + SysStringLen(bstr) as isize * SIZE_OF_U16 + 1 * SIZE_OF_U16 - 1)
    }; // ptr to last byte in buffer.
    debug_assert!(bstr_start_ptr < bstr_end_ptr); // empty BSTR anyway contains counter + 0x0000 marker in buffer;

    let src_start_ptr = &src[0] as *const u16 as *const u8;
    let src_end_ptr = unsafe { src_start_ptr.offset(src.len() as isize * SIZE_OF_U16 - 1) }; // ptr to last byte in slice
    debug_assert!(src_start_ptr < src_end_ptr); // src.len() > 0 a must.

    (bstr_start_ptr <= src_start_ptr && src_start_ptr <= bstr_end_ptr)
        || (src_start_ptr <= bstr_start_ptr && bstr_start_ptr <= src_end_ptr)
}

#[cfg(test)]
mod tests {
    use super::*;

    static TEST_LINE: &str = "Test line.\u{0000} Тестовая строка.\u{0000} Testlinie.\u{0000} Ligne de test.\u{0000} Línea de prueba.";

    fn bstr2string(bstr: BSTR) -> String {
        String::from_utf16_lossy(unsafe {
            std::slice::from_raw_parts(bstr, super::SysStringLen(bstr) as usize)
        })
    }

    #[test]
    fn test_SysAllocString() {
        // If successful, returns the string.
        let test_line_utf16: Vec<u16> = TEST_LINE.encode_utf16().collect();
        let bstr = SysAllocString(&test_line_utf16).unwrap();
        let original_sub_string = TEST_LINE.split("\u{0000}").take(1).next().unwrap();
        assert_eq!(original_sub_string, bstr2string(bstr));
        SysFreeString(bstr);

        // If source is a zero-length string, returns a zero-length BSTR.
        let zero_length = [0u16; 1];
        let bstr: BSTR = SysAllocString(&zero_length).unwrap();
        assert_eq!("", bstr2string(bstr));
        SysFreeString(bstr);

        // If source is not null-terminated, returns NullTerminatedStringRequiredError.
        let test_line_utf16: Vec<u16> = "Test line.".encode_utf16().collect();
        assert_eq!(
            Err(SysAllocError::NullTerminatedStringRequiredError),
            SysAllocString(&test_line_utf16)
        );
    }

    #[test]
    fn test_SysReAllocString() {
        // If successful, returns the string.
        let test_line_utf16: Vec<u16> = TEST_LINE.encode_utf16().collect();
        let bstr = SysAllocString(&test_line_utf16).unwrap();
        let new_test_line = "New line.\u{0000}";
        let test_line_utf16: Vec<u16> = new_test_line.encode_utf16().collect();
        let bstr = SysReAllocString(bstr, &test_line_utf16).unwrap();
        assert_eq!(new_test_line[..new_test_line.len() - 1], bstr2string(bstr));
        SysFreeString(bstr);

        // If source is a zero-length string, returns a zero-length BSTR.
        let test_line_utf16: Vec<u16> = TEST_LINE.encode_utf16().collect();
        let bstr = SysAllocString(&test_line_utf16).unwrap();
        let test_line_utf16 = [0u16, 1];
        let bstr = SysReAllocString(bstr, &test_line_utf16).unwrap();
        assert_eq!("", bstr2string(bstr));
        SysFreeString(bstr);

        // If bstr is NULL, returns InvalidPointerError.
        let bstr = NULL as BSTR;
        let test_line_utf16 = [0u16, 1];
        assert_eq!(
            Err(SysAllocError::InvalidPointerError),
            SysReAllocString(bstr, &test_line_utf16)
        );

        // If source is not null-terminated, returns NullTerminatedStringRequiredError.
        let test_line_utf16: Vec<u16> = TEST_LINE.encode_utf16().collect();
        let bstr = SysAllocString(&test_line_utf16).unwrap();
        let test_line_utf16: Vec<u16> = "New line.".encode_utf16().collect();
        assert_eq!(
            Err(SysAllocError::NullTerminatedStringRequiredError),
            SysReAllocString(bstr, &test_line_utf16)
        );
        SysFreeString(bstr);
    }

    #[test]
    fn test_SysAllocStringLen() {
        // If successful, returns the string.
        let test_line_utf16: Vec<u16> = TEST_LINE.encode_utf16().collect();
        let bstr = SysAllocStringLen(&test_line_utf16).unwrap();
        assert_eq!(TEST_LINE, bstr2string(bstr));
        SysFreeString(bstr);

        if std::usize::MAX > std::u32::MAX as usize {
            // If source is more than std::u32::MAX characters in length, returns SourceStringTooLongError.
            let bigfoot: Vec<u16> = vec![0; usize::try_from(std::u32::MAX).unwrap() + 1];
            assert_eq!(
                Err(SysAllocError::SourceStringTooLongError),
                SysAllocStringLen(&bigfoot)
            );
        }

        // If source is a zero-length string, returns a zero-length BSTR.
        let zero_length: Vec<u16> = vec![];
        let bstr: BSTR = SysAllocStringLen(&zero_length).unwrap();
        assert_eq!("", bstr2string(bstr));
        SysFreeString(bstr);
    }

    #[test]
    fn test_SysReAllocStringLen() {
        // If successful, returns the string.
        let test_line_utf16: Vec<u16> = TEST_LINE.encode_utf16().collect();
        let bstr = SysAllocStringLen(&test_line_utf16).unwrap();
        let new_test_line = "Test line.\u{0000} Тестовая строка.\u{0000} Testlinie.\u{0000} Ligne de test.\u{0000} Línea de prueba.\u{0000} 测试线";
        let test_line_utf16: Vec<u16> = new_test_line.encode_utf16().collect();
        let bstr = SysReAllocStringLen(bstr, &test_line_utf16).unwrap();
        assert_eq!(new_test_line, bstr2string(bstr));
        SysFreeString(bstr);

        // If source is a zero-length string, returns a zero-length BSTR.
        let test_line_utf16: Vec<u16> = TEST_LINE.encode_utf16().collect();
        let bstr = SysAllocStringLen(&test_line_utf16).unwrap();
        let test_line_utf16: Vec<u16> = vec![];
        let bstr = SysReAllocStringLen(bstr, &test_line_utf16).unwrap();
        assert_eq!("", bstr2string(bstr));
        SysFreeString(bstr);

        if std::usize::MAX > std::u32::MAX as usize {
            // If source is more than std::u32::MAX characters in length, returns SourceStringTooLongError.
            let test_line_utf16: Vec<u16> = TEST_LINE.encode_utf16().collect();
            let bstr = SysAllocStringLen(&test_line_utf16).unwrap();
            let bigfoot: Vec<u16> = vec![0; usize::try_from(std::u32::MAX).unwrap() + 1];
            assert_eq!(
                Err(SysAllocError::SourceStringTooLongError),
                SysAllocStringLen(&bigfoot)
            );
            SysFreeString(bstr);
        }

        // If bstr is NULL, returns InvalidPointerError.
        let bstr = NULL as BSTR;
        let test_line_utf16: Vec<u16> = vec![];
        assert_eq!(
            Err(SysAllocError::InvalidPointerError),
            SysReAllocStringLen(bstr, &test_line_utf16)
        );
    }

    #[test]
    fn test_SysStringLen() {
        let test_line_utf16: Vec<u16> = TEST_LINE.encode_utf16().collect();
        let bstr: BSTR = SysAllocStringLen(&test_line_utf16).unwrap();
        assert_eq!(test_line_utf16.len() as u32, SysStringLen(bstr));
        SysFreeString(bstr);

        let bstr: BSTR = NULL as BSTR;
        assert_eq!(0, SysStringLen(bstr));
    }

    #[test]
    fn test_bstr_src_intersection() {
        let test_line_utf16: Vec<u16> = TEST_LINE.encode_utf16().collect();
        let bstr = SysAllocStringLen(&test_line_utf16).unwrap();

        // src before bstr
        let src: &[u16] = unsafe { std::slice::from_raw_parts(bstr.offset(-4), 2) };
        assert!(!bstr_src_intersection(bstr, src));

        // src before bstr with intersection
        let src: &[u16] = unsafe { std::slice::from_raw_parts(bstr.offset(-4), 6) };
        assert!(bstr_src_intersection(bstr, src));

        // src = bstr (+ 32 counter and 0x0000 EOL marker)
        let src: &[u16] = unsafe {
            std::slice::from_raw_parts(bstr.offset(-2), 2 + SysStringLen(bstr) as usize + 1)
        };
        assert!(bstr_src_intersection(bstr, src));

        // src inside bstr
        let src: &[u16] = unsafe { std::slice::from_raw_parts(bstr.offset(2), 2) };
        assert!(bstr_src_intersection(bstr, src));

        // src after bstr with intersection
        let src: &[u16] =
            unsafe { std::slice::from_raw_parts(bstr.offset(SysStringLen(bstr) as isize - 1), 4) };
        assert!(bstr_src_intersection(bstr, src));

        // src after bstr
        let src: &[u16] =
            unsafe { std::slice::from_raw_parts(bstr.offset(SysStringLen(bstr) as isize + 1), 2) };
        assert!(!bstr_src_intersection(bstr, src));
    }
}
