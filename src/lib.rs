use anyhow::{anyhow, Result};
use std::{ptr, mem};
use uuid::Uuid;
use winapi::{
    ctypes::c_void, 
    shared::{
        guiddef::GUID,
        minwindef::DWORD,
    },
    um::{
        combaseapi::{CoCreateInstance, CoInitializeEx, CoUninitialize},
        oaidl::{IDispatch, DISPPARAMS, VARIANT},
        objbase::COINIT_APARTMENTTHREADED,
        winnt::HRESULT,
    },
};

///////////////////////////////////////////////////////////////////////////////
// Manually Define Missing Constants
///////////////////////////////////////////////////////////////////////////////

pub const CLSCTX_INPROC_SERVER: u32 = 0x1;
const S_OK: HRESULT = 0;
const DISPATCH_PROPERTYGET: u16 = 2; // ✅ Correct flag for COM property retrieval
const VT_I4: u16 = 3;
const LOCALE_USER_DEFAULT: u32 = 0x0400; 

// Standard IID_IDispatch
pub const IID_IDISPATCH: GUID = GUID {
    Data1: 0x00020400,
    Data2: 0x0000,
    Data3: 0x0000,
    Data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};

pub const GUID_NULL: GUID = GUID {
    Data1: 0,
    Data2: 0,
    Data3: 0,
    Data4: [0; 8],
};

///////////////////////////////////////////////////////////////////////////////
// QWQNG Wrapper Struct
///////////////////////////////////////////////////////////////////////////////

pub struct MedQrng {
    p_disp: *mut IDispatch,
}

impl MedQrng {
    pub fn new() -> Result<Self> {
        unsafe {
            // Initialize COM
            let hr = CoInitializeEx(ptr::null_mut(), COINIT_APARTMENTTHREADED);
            if hr != S_OK {
                return Err(anyhow!("CoInitializeEx failed: 0x{:08X}", hr));
            }

            // Hardcoded CLSID for QWQNG
            let clsid = get_clsid_from_registry()?;

            // Create COM object
            let mut p_disp: *mut IDispatch = ptr::null_mut();
            let hr = CoCreateInstance(
                &clsid,
                ptr::null_mut(),
                CLSCTX_INPROC_SERVER,
                &IID_IDISPATCH,
                &mut p_disp as *mut *mut IDispatch as *mut *mut c_void,
            );

            if hr != S_OK {
                CoUninitialize();
                return Err(anyhow!("CoCreateInstance failed: 0x{:08X}", hr));
            }

            Ok(Self { p_disp })
        }
    }

    /// Retrieves a 32-bit random integer from QWQNG (COM property)
    pub fn get_rand_int32(&self) -> Result<i32> {
        let variant = self.invoke_property("RandInt32")?;
        unsafe {
            let vt = variant.n1.n2().vt;
            if vt == VT_I4 {
                // `lVal()` returns a &i32, so we deref
                let val_ref = variant.n1.n2().n3.lVal();
                Ok(*val_ref)
            } else {
                Err(anyhow!(
                    "RandInt32 returned unexpected type: {} (expected VT_I4=3)",
                    vt
                ))
            }
        }
    }

    /// Calls a **COM property** using `DISPATCH_PROPERTYGET`
    fn invoke_property(&self, property_name: &str) -> Result<VARIANT> {
        unsafe {
            let dispid = get_dispid(self.p_disp, property_name)?;

            let mut params = DISPPARAMS {
                rgvarg: ptr::null_mut(),
                rgdispidNamedArgs: ptr::null_mut(),
                cArgs: 0,
                cNamedArgs: 0,
            };

            let mut var_result: VARIANT = mem::zeroed();
            let hr = (*self.p_disp).Invoke(
                dispid,
                &GUID_NULL,
                LOCALE_USER_DEFAULT,
                DISPATCH_PROPERTYGET, // ✅ Correct flag for retrieving COM properties
                &mut params,
                &mut var_result,
                ptr::null_mut(),
                ptr::null_mut(),
            );

            if hr == S_OK {
                Ok(var_result)
            } else {
                Err(anyhow!("Invoke('{}') failed: 0x{:08X}", property_name, hr))
            }
        }
    }
}

// Uninitialize COM when the struct is dropped
impl Drop for MedQrng {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

///////////////////////////////////////////////////////////////////////////////
// Hardcoded CLSID for QWQNG
///////////////////////////////////////////////////////////////////////////////

fn get_clsid_from_registry() -> Result<GUID> {
    // The actual CLSID you found in the registry:
    let clsid_str = "{D7A1BFCF-9A30-45AF-A5E4-2CAF0A344938}";
    let uuid = Uuid::parse_str(clsid_str.trim())?;
    Ok(uuid_to_winapi_guid(&uuid))
}

/// Convert a `uuid::Uuid` to a `winapi::GUID`
fn uuid_to_winapi_guid(u: &Uuid) -> GUID {
    let b = u.as_bytes();
    GUID {
        Data1: u32::from_be_bytes([b[0], b[1], b[2], b[3]]),
        Data2: u16::from_be_bytes([b[4], b[5]]),
        Data3: u16::from_be_bytes([b[6], b[7]]),
        Data4: [b[8], b[9], b[10], b[11], b[12], b[13], b[14], b[15]],
    }
}

///////////////////////////////////////////////////////////////////////////////
// Get DISP ID for method invocation
///////////////////////////////////////////////////////////////////////////////

unsafe fn get_dispid(p_disp: *mut IDispatch, name: &str) -> Result<i32> {
    let wide_name = to_utf16(name);
    let mut dispid = 0i32;
    let mut rgsz_names = [wide_name.as_ptr() as *mut u16];

    let hr = (*p_disp).GetIDsOfNames(
        &GUID_NULL,
        rgsz_names.as_mut_ptr(),
        1,
        LOCALE_USER_DEFAULT,
        &mut dispid,
    );

    if hr == S_OK {
        Ok(dispid)
    } else {
        Err(anyhow!("GetIDsOfNames('{}') failed: 0x{:08X}", name, hr))
    }
}

fn to_utf16(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}
