fn main() -> anyhow::Result<()> {
    let qrng = med_qrng::MedQrng::new()?;

    // Test RandInt32
    let i32_val = qrng.rand_int32()?;
    println!("RandInt32 = {}", i32_val);

    // Test RandUniform
    let uniform_val = qrng.rand_uniform()?;
    println!("RandUniform = {}", uniform_val);

    // Test RandNormal
    let normal_val = qrng.rand_normal()?;
    println!("RandNormal = {}", normal_val);

    // Test RandBytes (request 10 bytes)
    let bytes_val = qrng.rand_bytes(10)?;
    println!("RandBytes(10) = {:?}", bytes_val);

    // Test DeviceId
    let dev_id = qrng.device_id()?;
    println!("DeviceId = {}", dev_id);

    // Test RuntimeInfo (SAFEARRAY of f32)
    let rt_info = qrng.runtime_info()?;
    println!("RuntimeInfo = {:?}", rt_info);

    // Test Diagnostics (example: dx code 0x15)
    let dx_data = qrng.diagnostics(0x15)?;
    println!("Diagnostics(0x15) = {:?}", dx_data);

    // Test Clear and Reset methods
    qrng.clear()?;
    println!("Buffers cleared.");
    qrng.reset()?;
    println!("Device reset.");

    Ok(())
}
