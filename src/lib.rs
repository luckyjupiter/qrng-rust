use anyhow::{anyhow, Result};
use std::{mem, ptr};
use uuid::Uuid;
use winapi::{
    ctypes::c_void,
    shared::guiddef::GUID,
    um::{
        combaseapi::{CoCreateInstance, CoInitializeEx, CoUninitialize},
        cguid::GUID_NULL, // Import GUID_NULL
        oleauto::{SafeArrayAccessData, SafeArrayGetLBound, SafeArrayGetUBound, SafeArrayUnaccessData, SysStringLen},
        oaidl::{IDispatch, DISPPARAMS, VARIANT},
        oaidl::SAFEARRAY, // Import SAFEARRAY from the public module
        objbase::COINIT_APARTMENTTHREADED,
        winnt::HRESULT,
    },
};

///////////////////////////////////////////////////////////////////////////////
// Constants for COM and VARIANT types
///////////////////////////////////////////////////////////////////////////////

pub const CLSCTX_INPROC_SERVER: u32 = 0x1;
const S_OK: HRESULT = 0;
const DISPATCH_PROPERTYGET: u16 = 2;
const DISPATCH_METHOD: u16 = 1;

const VT_I4: u16 = 3;    // 32-bit integer
const VT_R8: u16 = 5;    // Double (f64)
const VT_R4: u16 = 4;    // 32-bit float (f32)
const VT_BSTR: u16 = 8;  // BSTR (wide string)
const VT_ARRAY: u16 = 0x2000; // Flag indicating SAFEARRAY
const VT_UI1: u16 = 17;  // Unsigned 8-bit integer

const LOCALE_USER_DEFAULT: u32 = 0x0400;

pub const IID_IDISPATCH: GUID = GUID {
    Data1: 0x00020400,
    Data2: 0x0000,
    Data3: 0x0000,
    Data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};

///////////////////////////////////////////////////////////////////////////////
// QWQNG Library Struct
///////////////////////////////////////////////////////////////////////////////

pub struct MedQrng {
    p_disp: *mut IDispatch,
}

impl MedQrng {
    /// Creates and initializes the QWQNG COM object.
    pub fn new() -> Result<Self> {
        unsafe {
            let hr = CoInitializeEx(ptr::null_mut(), COINIT_APARTMENTTHREADED);
            if hr != S_OK {
                return Err(anyhow!("CoInitializeEx failed: 0x{:08X}", hr));
            }
            let clsid = get_qwqng_clsid()?;
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

    /// Retrieves a 32-bit random integer from the RandInt32 property.
    pub fn rand_int32(&self) -> Result<i32> {
        let var = self.invoke_property("RandInt32", &[])?;
        unsafe {
            if var.n1.n2().vt == VT_I4 {
                Ok(*var.n1.n2().n3.lVal())
            } else {
                Err(anyhow!("RandInt32 returned non-i32 type"))
            }
        }
    }

    /// Retrieves a uniform random double (in [0,1)) from RandUniform.
    pub fn rand_uniform(&self) -> Result<f64> {
        let var = self.invoke_property("RandUniform", &[])?;
        unsafe {
            if var.n1.n2().vt == VT_R8 {
                Ok(*var.n1.n2().n3.dblVal())
            } else {
                Err(anyhow!("RandUniform returned non-f64 type"))
            }
        }
    }

    /// Retrieves a normally distributed random double from RandNormal.
    pub fn rand_normal(&self) -> Result<f64> {
        let var = self.invoke_property("RandNormal", &[])?;
        unsafe {
            if var.n1.n2().vt == VT_R8 {
                Ok(*var.n1.n2().n3.dblVal())
            } else {
                Err(anyhow!("RandNormal returned non-f64 type"))
            }
        }
    }

    /// Retrieves random bytes (SAFEARRAY of VT_UI1) from RandBytes.
    /// Pass the desired byte length as an argument.
    pub fn rand_bytes(&self, length: i32) -> Result<Vec<u8>> {
        let var = self.invoke_property_with_i32_arg("RandBytes", length)?;
        variant_to_byte_array(&var)
    }

    /// Retrieves the device serial number (BSTR) from DeviceId.
    pub fn device_id(&self) -> Result<String> {
        let var = self.invoke_property("DeviceId", &[])?;
        variant_to_bstr(&var)
    }

    /// Retrieves runtime info (SAFEARRAY of VT_R4) from RuntimeInfo.
    pub fn runtime_info(&self) -> Result<Vec<f32>> {
        let var = self.invoke_property("RuntimeInfo", &[])?;
        variant_to_f32_array(&var)
    }

    /// Retrieves diagnostics data (SAFEARRAY of VT_UI1) from Diagnostics.
    /// In our implementation Diagnostics is invoked as a method.
    pub fn diagnostics(&self, dx_code: i32) -> Result<Vec<u8>> {
        let var = self.invoke_method_with_i32_arg("Diagnostics", dx_code)?;
        variant_to_byte_array(&var)
    }

    /// Calls the Clear() method.
    pub fn clear(&self) -> Result<()> {
        self.invoke_method("Clear", &[])?;
        Ok(())
    }

    /// Calls the Reset() method.
    pub fn reset(&self) -> Result<()> {
        self.invoke_method("Reset", &[])?;
        Ok(())
    }
}

impl Drop for MedQrng {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

///////////////////////////////////////////////////////////////////////////////
// Helper: Invoking COM Properties and Methods
///////////////////////////////////////////////////////////////////////////////

impl MedQrng {
    /// Invokes a COM property (DISPATCH_PROPERTYGET) with optional arguments.
    fn invoke_property(&self, name: &str, args: &[VARIANT]) -> Result<VARIANT> {
        unsafe {
            let dispid = get_dispid(self.p_disp, name)?;
            let mut dp: DISPPARAMS = mem::zeroed();
            if !args.is_empty() {
                dp.rgvarg = args.as_ptr() as *mut VARIANT;
                dp.cArgs = args.len() as u32;
            }
            let mut var_result: VARIANT = mem::zeroed();
            let hr = (*self.p_disp).Invoke(
                dispid,
                &GUID_NULL,
                LOCALE_USER_DEFAULT,
                DISPATCH_PROPERTYGET,
                &mut dp,
                &mut var_result,
                ptr::null_mut(),
                ptr::null_mut(),
            );
            if hr == S_OK {
                Ok(var_result)
            } else {
                Err(anyhow!("Invoke('{}') failed: 0x{:08X}", name, hr))
            }
        }
    }

    /// Invokes a COM property with a single i32 argument.
    fn invoke_property_with_i32_arg(&self, name: &str, arg: i32) -> Result<VARIANT> {
        let mut var_arg: VARIANT = unsafe { mem::zeroed() };
        unsafe {
            var_arg.n1.n2_mut().vt = VT_I4;
            *var_arg.n1.n2_mut().n3.lVal_mut() = arg;
        }
        self.invoke_property(name, &[var_arg])
    }

    /// Invokes a COM method (DISPATCH_METHOD) with optional arguments.
    /// This version does not expect a return value.
    fn invoke_method(&self, name: &str, args: &[VARIANT]) -> Result<()> {
        unsafe {
            let _ = self.invoke_method_return(name, args)?;
            Ok(())
        }
    }

    /// Invokes a COM method (DISPATCH_METHOD) with optional arguments and returns the VARIANT.
    fn invoke_method_return(&self, name: &str, args: &[VARIANT]) -> Result<VARIANT> {
        unsafe {
            let dispid = get_dispid(self.p_disp, name)?;
            let mut dp: DISPPARAMS = mem::zeroed();
            if !args.is_empty() {
                dp.rgvarg = args.as_ptr() as *mut VARIANT;
                dp.cArgs = args.len() as u32;
            }
            let mut var_result: VARIANT = mem::zeroed();
            let hr = (*self.p_disp).Invoke(
                dispid,
                &GUID_NULL,
                LOCALE_USER_DEFAULT,
                DISPATCH_METHOD,
                &mut dp,
                &mut var_result,
                ptr::null_mut(),
                ptr::null_mut(),
            );
            if hr == S_OK {
                Ok(var_result)
            } else {
                Err(anyhow!("Invoke('{}') failed: 0x{:08X}", name, hr))
            }
        }
    }

    /// Invokes a COM method with a single i32 argument and returns the VARIANT.
    fn invoke_method_with_i32_arg(&self, name: &str, arg: i32) -> Result<VARIANT> {
        let mut var_arg: VARIANT = unsafe { mem::zeroed() };
        unsafe {
            var_arg.n1.n2_mut().vt = VT_I4;
            *var_arg.n1.n2_mut().n3.lVal_mut() = arg;
        }
        self.invoke_method_return(name, &[var_arg])
    }
}

///////////////////////////////////////////////////////////////////////////////
// Helper: Get QWQNG CLSID (hardcoded)
///////////////////////////////////////////////////////////////////////////////

fn get_qwqng_clsid() -> Result<GUID> {
    let clsid_str = "{D7A1BFCF-9A30-45AF-A5E4-2CAF0A344938}";
    let uuid = Uuid::parse_str(clsid_str.trim())?;
    Ok(uuid_to_winapi_guid(&uuid))
}

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
// Helper: Get DISP ID for a COM member
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

///////////////////////////////////////////////////////////////////////////////
// SAFEARRAY / BSTR Parsing Functions
///////////////////////////////////////////////////////////////////////////////

fn variant_to_byte_array(var: &VARIANT) -> Result<Vec<u8>> {
    unsafe {
        let vt = var.n1.n2().vt;
        if (vt & VT_ARRAY) != VT_ARRAY || (vt & VT_UI1) != VT_UI1 {
            return Err(anyhow!("Expected SAFEARRAY of bytes, but got vt=0x{:X}", vt));
        }
        // Get the SAFEARRAY pointer by calling parray() and dereferencing
        let psa: *mut SAFEARRAY = *var.n1.n2().n3.parray();
        if psa.is_null() {
            return Err(anyhow!("Null SAFEARRAY pointer"));
        }
        let mut lbound: i32 = 0;
        let mut ubound: i32 = 0;
        let hr_lb = SafeArrayGetLBound(psa, 1, &mut lbound as *mut i32);
        let hr_ub = SafeArrayGetUBound(psa, 1, &mut ubound as *mut i32);
        if hr_lb != S_OK || hr_ub != S_OK {
            return Err(anyhow!("SafeArrayGetLBound/UBound failed"));
        }
        let count = (ubound - lbound + 1) as usize;
        let mut data_ptr: *mut u8 = ptr::null_mut();
        let hr_access = SafeArrayAccessData(psa, &mut data_ptr as *mut *mut u8 as *mut *mut c_void);
        if hr_access != S_OK {
            return Err(anyhow!("SafeArrayAccessData failed"));
        }
        let slice = std::slice::from_raw_parts(data_ptr, count);
        let bytes = slice.to_vec();
        SafeArrayUnaccessData(psa);
        Ok(bytes)
    }
}

fn variant_to_bstr(var: &VARIANT) -> Result<String> {
    unsafe {
        let vt = var.n1.n2().vt;
        if vt != VT_BSTR {
            return Err(anyhow!("Expected BSTR, but got vt=0x{:X}", vt));
        }
        let bstr_ptr = *var.n1.n2().n3.bstrVal();
        if bstr_ptr.is_null() {
            return Ok(String::new());
        }
        let len = SysStringLen(bstr_ptr) as usize;
        let slice = std::slice::from_raw_parts(bstr_ptr, len);
        let rust_string = String::from_utf16_lossy(slice);
        Ok(rust_string)
    }
}

fn variant_to_f32_array(var: &VARIANT) -> Result<Vec<f32>> {
    unsafe {
        let vt = var.n1.n2().vt;
        if (vt & VT_ARRAY) != VT_ARRAY || (vt & VT_R4) != VT_R4 {
            return Err(anyhow!("Expected SAFEARRAY of f32 (VT_ARRAY|VT_R4), got vt=0x{:X}", vt));
        }
        let psa: *mut SAFEARRAY = *var.n1.n2().n3.parray();
        if psa.is_null() {
            return Err(anyhow!("Null SAFEARRAY pointer for float array"));
        }
        let mut lbound: i32 = 0;
        let mut ubound: i32 = 0;
        if SafeArrayGetLBound(psa, 1, &mut lbound as *mut i32) != S_OK {
            return Err(anyhow!("SafeArrayGetLBound failed"));
        }
        if SafeArrayGetUBound(psa, 1, &mut ubound as *mut i32) != S_OK {
            return Err(anyhow!("SafeArrayGetUBound failed"));
        }
        let count = (ubound - lbound + 1) as usize;
        let mut data_ptr: *mut f32 = ptr::null_mut();
        if SafeArrayAccessData(psa, &mut data_ptr as *mut *mut f32 as *mut *mut c_void) != S_OK {
            return Err(anyhow!("SafeArrayAccessData failed on float array"));
        }
        let slice = std::slice::from_raw_parts(data_ptr, count);
        let floats = slice.to_vec();
        SafeArrayUnaccessData(psa);
        Ok(floats)
    }
}
