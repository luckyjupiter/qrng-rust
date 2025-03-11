fn main() -> anyhow::Result<()> {
    let qrng = med_qrng::MedQrng::new()?;
    let rand_int = qrng.get_rand_int32()?;
    println!("Random int from QWQNG: {}", rand_int);
    Ok(())
}
