#![allow(non_camel_case_types, non_snake_case, unused)]

//! Smart & safe rustified WinAPI IDispatch counterpart.
//!

use std::cell::Cell;
use std::convert::{TryFrom, TryInto};
use std::error::Error;

use winapi::shared::guiddef::{IID_NULL, REFIID};
use winapi::shared::minwindef::{LPVOID, PUINT, UINT, WORD};
use winapi::shared::ntdef::{HRESULT, INT, LCID, PULONG, ULONG};
use winapi::shared::winerror;
use winapi::shared::wtypes::{BSTR, DATE, VARIANT_BOOL};
use winapi::shared::wtypesbase::LPOLESTR;
use winapi::um::oaidl::{
    IDispatch, IDispatchVtbl, ITypeInfo, DISPID, DISPID_NEWENUM, DISPPARAMS, EXCEPINFO, LPDISPATCH,
    LPVARIANT, SAFEARRAY, VARIANT,
};
use winapi::um::oleauto::{
    SysStringLen, VariantClear, VariantInit, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
    DISPATCH_PROPERTYPUT,
};
use winapi::um::unknwnbase::{IClassFactory, IClassFactoryVtbl, IUnknown, IUnknownVtbl, LPUNKNOWN};
use winapi::um::winnt::{LOCALE_USER_DEFAULT, LONG, LPCSTR, LPSTR, WCHAR};
use winapi::{Class, Interface, RIDL};

use crate::auto_bstr::*;
use crate::auto_com_interface::*;
use crate::smart_iunknown::*;
use crate::smart_variant::*;

pub trait SmartIDispatch: SmartIUnknown {
    fn as_idispatch(&self) -> &IDispatch;
    fn as_idispatch_mut(&mut self) -> &mut IDispatch;

    fn get_type_info_count(&self) -> Result<UINT, HRESULT> {
        let mut pctinfo: UINT = 0;
        let hresult = unsafe { self.as_idispatch().GetTypeInfoCount(&mut pctinfo) };
        if winerror::SUCCEEDED(hresult) {
            Ok(pctinfo)
        } else {
            Err(hresult)
        }
    }

    fn get_type_info(
        &self,
        iTInfo: UINT,
        lcid: LCID,
    ) -> Result<AutoCOMInterface<ITypeInfo>, HRESULT> {
        let mut ptinfo: *mut ITypeInfo = std::ptr::null_mut();
        let hresult = unsafe { self.as_idispatch().GetTypeInfo(iTInfo, lcid, &mut ptinfo) };
        if winerror::SUCCEEDED(hresult) {
            unsafe { Ok((ptinfo as *mut ITypeInfo).try_into().unwrap()) }
        } else {
            Err(hresult)
        }
    }

    fn get_ids_of_names(&self, names: &[&str], lcid: LCID) -> (Vec<DISPID>, HRESULT) {
        let cNames: UINT = names.len() as UINT;
        let mut rgDispId: Vec<DISPID> = vec![-1; cNames as usize];
        let mut szNames: Vec<Vec<u16>> = names
            .iter()
            .map(|x| x.encode_utf16().chain(std::iter::once(0)).collect())
            .collect();
        let mut rgszNames: Vec<LPOLESTR> = szNames.iter_mut().map(|x| x.as_mut_ptr()).collect();

        let hresult = unsafe {
            self.as_idispatch().GetIDsOfNames(
                &IID_NULL,
                rgszNames.as_mut_ptr(),
                cNames,
                lcid,
                rgDispId.as_mut_ptr(),
            )
        };

        (rgDispId, hresult)
    }

    fn invoke(
        &mut self,
        member_dispid: DISPID,
        lcid: LCID,
        flags: WORD,
        params: &[SmartVariant],
    ) -> Result<SmartVariant, (HRESULT, String, u32)> {
        let mut rev_params: Vec<VARIANT> = params.iter().cloned().map(|x| x.into()).rev().collect();
        let mut result = VARIANT::default();

        unsafe {
            let mut dispparams = DISPPARAMS {
                cArgs: rev_params.len() as u32,
                rgvarg: rev_params.as_mut_ptr(),
                rgdispidNamedArgs: std::ptr::null_mut() as *mut DISPID,
                cNamedArgs: 0,
            };

            let mut ex_info: EXCEPINFO = std::mem::zeroed();
            let mut arg = UINT::default();

            let hresult = self.as_idispatch_mut().Invoke(
                member_dispid,
                &IID_NULL,
                lcid,
                flags,
                &mut dispparams,
                &mut result,
                &mut ex_info,
                &mut arg,
            );

            if winapi::shared::winerror::SUCCEEDED(hresult) {
                Ok(result.into())
            } else {
                Err((hresult, AutoBSTR::from(ex_info.bstrDescription).into(), arg))
            }
        }
    }

    fn call(
        &mut self,
        method: &str,
        params: &[SmartVariant],
    ) -> Result<SmartVariant, (HRESULT, String, u32)> {
        match self.get_ids_of_names(&[method], LOCALE_USER_DEFAULT) {
            (ids, S_OK) => self.invoke(ids[0], LOCALE_USER_DEFAULT, DISPATCH_METHOD, params),
            (_, e) => Err((e, "get_ids_of_names()".into(), 0)),
        }
    }

    fn get(&mut self, property: &str) -> Result<SmartVariant, (HRESULT, String, u32)> {
        match self.get_ids_of_names(&[property], LOCALE_USER_DEFAULT) {
            (ids, S_OK) => self.invoke(ids[0], LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &[]),
            (_, e) => Err((e, "get_ids_of_names()".into(), 0)),
        }
    }

    fn put(
        &mut self,
        property: &str,
        value: SmartVariant,
    ) -> Result<SmartVariant, (HRESULT, String, u32)> {
        match self.get_ids_of_names(&[property], LOCALE_USER_DEFAULT) {
            (ids, S_OK) => self.invoke(ids[0], LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &[value]),
            (_, e) => Err((e, "get_ids_of_names()".into(), 0)),
        }
    }
}

impl SmartIDispatch for IDispatch {
    fn as_idispatch(&self) -> &IDispatch {
        self
    }

    fn as_idispatch_mut(&mut self) -> &mut IDispatch {
        self
    }
}

impl SmartIDispatch for AutoCOMInterface<IDispatch> {
    fn as_idispatch(&self) -> &IDispatch {
        self.as_inner()
    }

    fn as_idispatch_mut(&mut self) -> &mut IDispatch {
        self.as_inner_mut()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::auto_bstr::*;
    use std::convert::TryInto;
    use winapi::um::combaseapi::{CoCreateInstance, CoGetClassObject, CLSCTX_ALL};

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

    #[test]
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

        let mut conn1Cdb: AutoCOMInterface<IDispatch> = conn1Cdb.try_into().unwrap();

        let dispids = conn1Cdb.get_ids_of_names(
            &[
                "NewObject",
                "ПолучитьСтруктуруХраненияБазыДанных",
            ],
            LOCALE_USER_DEFAULT,
        );

        assert!(winapi::shared::winerror::SUCCEEDED(dispids.1));

        assert_eq!(dispids.0[1], 0);

        // let mut kv: AutoCOMInterface<IDispatch> = conn1Cdb
        //     .call(
        //         "ПолучитьСтруктуруХраненияБазыДанных",
        //         &[SmartVariant::Empty, SmartVariant::Bool(true)],
        //     )
        //     .unwrap()
        //     .try_into()
        //     .unwrap();
        // let count = kv.call("Количество", &[]).unwrap();

        // assert_eq!(count, SmartVariant::Int4(0));

        unsafe { winapi::um::combaseapi::CoUninitialize() };
    }
}
