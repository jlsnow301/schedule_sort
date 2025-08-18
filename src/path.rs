use std::{fs, path::Path};

use anyhow::{anyhow, Context, Result};

const MAX_FILE_SIZE: u64 = 10 * 1024 * 1024;

pub fn get_file_path(arg: &str) -> Result<String> {
    if arg.is_empty() {
        return Err(anyhow!("No file input"));
    }

    let path = Path::new(arg);

    if !path.exists() {
        return Err(anyhow!("File doesn't exist"));
    }

    if !path.is_file() {
        return Err(anyhow!("Path isn't a file"));
    }

    let file_type = path
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_lowercase());

    if file_type.as_deref().unwrap() != "xlsx" {
        return Err(anyhow!("File isn't of .xlsx format"));
    }

    let metadata = fs::metadata(path).context("Could not read metadata")?;
    let file_size = metadata.len();
    let size_mb = file_size as f64 / (1024.0 * 1024.0);
    if file_size > MAX_FILE_SIZE {
        return Err(anyhow!("File ({}) is over max size", size_mb));
    }

    Ok(path.to_string_lossy().to_string())
}
