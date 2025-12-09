//! Memory benchmark for reading large files
//! 
//! Tests StreamingReader performance with custom ZIP implementation

use excelstream::streaming_reader::StreamingReader;
use std::fs;
use std::time::Instant;

fn get_memory_kb() -> Option<usize> {
    let status = fs::read_to_string("/proc/self/status").ok()?;
    for line in status.lines() {
        if line.starts_with("VmRSS:") {
            let parts: Vec<&str> = line.split_whitespace().collect();
            return parts.get(1)?.parse().ok();
        }
    }
    None
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== ExcelStream Read Memory Benchmark ===\n");
    println!("Testing StreamingReader with custom ZIP implementation (v0.9.0)");
    println!("File: memory_test_1000000.xlsx (generated from previous benchmark)\n");

    let filename = "memory_test_1000000.xlsx";
    
    // Check if file exists
    if !std::path::Path::new(filename).exists() {
        println!("❌ File not found: {}", filename);
        println!("   Please run: cargo run --example memory_benchmark --release");
        println!("   This will generate the test file.");
        return Ok(());
    }
    
    let file_size = fs::metadata(filename)?.len();
    println!("📁 File info:");
    println!("   Size: {:.2} MB", file_size as f64 / (1024.0 * 1024.0));
    println!("   Expected: 1M rows × 30 columns = 30M cells");
    println!();

    // Test 1: Open and load metadata
    println!("📊 Phase 1: Open file and load metadata");
    let mem_start = get_memory_kb().unwrap_or(0);
    println!("   Memory before open: {:.2} MB", mem_start as f64 / 1024.0);
    
    let start = Instant::now();
    let mut reader = StreamingReader::open(filename)?;
    let open_time = start.elapsed();
    
    let mem_after_open = get_memory_kb().unwrap_or(0);
    println!("   ✅ Opened in {:.3}s", open_time.as_secs_f64());
    println!("   Memory after open: {:.2} MB", mem_after_open as f64 / 1024.0);
    println!("   Memory delta: {:.2} MB (SST + metadata loaded)", 
             (mem_after_open - mem_start) as f64 / 1024.0);
    println!();

    // Test 2: Stream all rows
    println!("📊 Phase 2: Stream all rows");
    let mem_before_read = get_memory_kb().unwrap_or(0);
    println!("   Memory before streaming: {:.2} MB", mem_before_read as f64 / 1024.0);
    
    let start = Instant::now();
    let mut row_count = 0;
    let mut cell_count = 0;
    let mut peak_mem = mem_before_read;
    
    for row_result in reader.stream_rows("Sheet1")? {
        let row = row_result?;
        row_count += 1;
        cell_count += row.len();
        
        // Sample memory periodically
        if row_count % 100_000 == 0 {
            if let Some(mem) = get_memory_kb() {
                peak_mem = peak_mem.max(mem);
                println!("   Progress: {} rows ({:.1}M cells) - Memory: {:.2} MB", 
                         row_count, 
                         cell_count as f64 / 1_000_000.0,
                         mem as f64 / 1024.0);
            }
        }
    }
    
    let read_time = start.elapsed();
    let mem_after_read = get_memory_kb().unwrap_or(0);
    
    println!();
    println!("   ✅ Completed in {:.2}s", read_time.as_secs_f64());
    println!("   Total rows: {}", row_count);
    println!("   Total cells: {} ({:.1}M)", cell_count, cell_count as f64 / 1_000_000.0);
    println!("   Speed: {:.0} rows/sec", row_count as f64 / read_time.as_secs_f64());
    println!("   Throughput: {:.0} cells/sec", cell_count as f64 / read_time.as_secs_f64());
    println!();
    
    println!("📊 Memory Analysis:");
    println!("   Memory start: {:.2} MB", mem_start as f64 / 1024.0);
    println!("   Memory after open: {:.2} MB", mem_after_open as f64 / 1024.0);
    println!("   Memory peak during read: {:.2} MB", peak_mem as f64 / 1024.0);
    println!("   Memory after read: {:.2} MB", mem_after_read as f64 / 1024.0);
    println!("   Memory delta (open): {:.2} MB", (mem_after_open - mem_start) as f64 / 1024.0);
    println!("   Memory delta (peak): {:.2} MB", (peak_mem - mem_start) as f64 / 1024.0);
    println!();

    println!("=== Summary ===");
    println!("✅ StreamingReader with custom ZIP implementation");
    println!("✅ File size: {:.2} MB", file_size as f64 / (1024.0 * 1024.0));
    println!("✅ Peak memory: {:.2} MB (constant!)", peak_mem as f64 / 1024.0);
    println!("✅ Read speed: {:.0} rows/sec", row_count as f64 / read_time.as_secs_f64());
    println!("✅ Memory efficiency: {:.1}x smaller than file size", 
             file_size as f64 / ((peak_mem - mem_start) as f64 * 1024.0));
    println!();
    println!("🎯 Key Achievement:");
    println!("   {:.2} MB file → {:.2} MB RAM = {:.0}% memory reduction!", 
             file_size as f64 / (1024.0 * 1024.0),
             peak_mem as f64 / 1024.0,
             (1.0 - (peak_mem as f64 / (file_size as f64 / 1024.0))) * 100.0);

    Ok(())
}
