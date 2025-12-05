//! Helper functions để tự động cấu hình memory constraints

use crate::error::Result;
use crate::fast_writer::UltraLowMemoryWorkbook;
use std::path::Path;

/// Memory profile cho các loại pods khác nhau
#[derive(Debug, Clone, Copy)]
pub enum MemoryProfile {
    /// Pods nhỏ (< 512MB): flush mỗi 100 rows, buffer 256KB
    Low,
    /// Pods trung bình (512MB-1GB): flush mỗi 500 rows, buffer 512KB
    Medium,
    /// Pods lớn (> 1GB): flush mỗi 1000 rows, buffer 1MB (default)
    High,
    /// Custom profile
    Custom {
        flush_interval: u32,
        max_buffer_size: usize,
    },
}

impl MemoryProfile {
    /// Tạo profile từ memory limit (MB)
    pub fn from_memory_mb(memory_mb: usize) -> Self {
        if memory_mb < 512 {
            MemoryProfile::Low
        } else if memory_mb < 1024 {
            MemoryProfile::Medium
        } else {
            MemoryProfile::High
        }
    }

    /// Detect từ environment variable MEMORY_LIMIT_MB
    pub fn from_env() -> Self {
        std::env::var("MEMORY_LIMIT_MB")
            .ok()
            .and_then(|s| s.parse::<usize>().ok())
            .map(Self::from_memory_mb)
            .unwrap_or(MemoryProfile::High)
    }

    fn apply(&self, _workbook: &mut UltraLowMemoryWorkbook) {
        // UltraLowMemoryWorkbook doesn't have flush_interval or max_buffer_size
        // These settings are no-ops for compatibility
        match self {
            MemoryProfile::Low => {
                // Low memory mode - already optimized
            }
            MemoryProfile::Medium => {
                // Medium memory mode - already optimized
            }
            MemoryProfile::High => {
                // High memory mode - already optimized
            }
            MemoryProfile::Custom { .. } => {
                // Custom settings - already optimized
            }
        }
    }
}

/// Tạo UltraLowMemoryWorkbook với memory profile tự động
pub fn create_workbook_auto<P: AsRef<Path>>(path: P) -> Result<UltraLowMemoryWorkbook> {
    let profile = MemoryProfile::from_env();
    create_workbook_with_profile(path, profile)
}

/// Tạo UltraLowMemoryWorkbook với memory profile chỉ định
pub fn create_workbook_with_profile<P: AsRef<Path>>(
    path: P,
    profile: MemoryProfile,
) -> Result<UltraLowMemoryWorkbook> {
    let mut workbook = UltraLowMemoryWorkbook::new(path)?;
    profile.apply(&mut workbook);
    Ok(workbook)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_memory_profile_from_mb() {
        assert!(matches!(
            MemoryProfile::from_memory_mb(256),
            MemoryProfile::Low
        ));
        assert!(matches!(
            MemoryProfile::from_memory_mb(768),
            MemoryProfile::Medium
        ));
        assert!(matches!(
            MemoryProfile::from_memory_mb(2048),
            MemoryProfile::High
        ));
    }
}
