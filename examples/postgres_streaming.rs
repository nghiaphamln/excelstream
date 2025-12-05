//! Memory-efficient PostgreSQL to Excel streaming export
//!
//! This example demonstrates the most memory-efficient way to export
//! millions of rows from PostgreSQL to Excel with minimal memory footprint.
//!
//! Key features:
//! - Uses PostgreSQL cursor for server-side result streaming
//! - Processes data in small batches
//! - Minimal memory usage (suitable for 10M+ rows)
//! - Progress tracking with ETA
//! - Uses typed values for better performance (+40% faster)

use excelstream::types::CellValue;
use excelstream::writer::ExcelWriter;
use postgres::{Client, NoTls};
use std::time::{Duration, Instant};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Memory-Efficient PostgreSQL to Excel Export ===\n");

    // Configuration
    let connection_string = std::env::var("DATABASE_URL").unwrap_or_else(|_| {
        "postgres://u_e_invoice:abcd%401234@10.98.98.100:5432/e_invoice".to_string()
    });

    let output_file = "e_invoice_export.xlsx";
    // MEMORY OPTIMIZATION: Reduced batch size from 5000 to 500 rows
    // Larger batches = more memory pressure. 500 rows = ~1-2 MB per batch
    let batch_size: i32 = 500;

    // SQL query - modify as needed
    let query = r#"
        SELECT id, order_numbers, warehouse_code, invoice_no, invoice_pattern, invoice_serial,
               fpt_transaction_id, document_date, issue_date, status,
               buyer_name, buyer_tax_code, buyer_address, buyer_email, buyer_budget_relation_id,
               tax_authority_code, total_amount, vat_amount, grand_total,
               last_sync_attempt_at, last_sync_error_message, secret_code, signing_date,
               created_at, updated_at
        FROM e_invoices
        WHERE document_date >= '2025-09-28' AND document_date <= '2025-12-05'
        ORDER BY document_date DESC
    "#;

    println!("Configuration:");
    println!("  Output file: {}", output_file);
    println!("  Batch size: {} (memory-optimized)", batch_size);
    println!("  Query: {}\n", query.trim());

    let start = Instant::now();

    // Connect to PostgreSQL
    println!("Connecting to PostgreSQL...");
    let mut client = Client::connect(&connection_string, NoTls)?;
    println!("Connected.\n");

    // Start a transaction to use cursors
    let mut transaction = client.transaction()?;

    // Declare a server-side cursor
    println!("Creating server-side cursor...");
    transaction.execute(
        r#"
        DECLARE export_cursor CURSOR FOR
        SELECT id, order_numbers, warehouse_code, invoice_no, invoice_pattern, invoice_serial,
               fpt_transaction_id, document_date, issue_date, status,
               buyer_name, buyer_tax_code, buyer_address, buyer_email, buyer_budget_relation_id,
               tax_authority_code, total_amount, vat_amount, grand_total,
               last_sync_attempt_at, last_sync_error_message, secret_code, signing_date,
               created_at, updated_at
        FROM e_invoices
        WHERE document_date >= '2025-09-28' AND document_date <= '2025-12-05'
        ORDER BY document_date DESC
        "#,
        &[],
    )?;

    // Create Excel workbook
    println!("Creating Excel workbook...");
    let mut writer = ExcelWriter::new(output_file)?;

    // Configure for optimal memory usage
    writer.set_flush_interval(500); // Flush every 500 rows (lower = more frequent flushing = lower peak memory)
    writer.set_max_buffer_size(512 * 1024); // Reduced from 1MB to 512KB to force more frequent flushes

    // Write header for e_invoice data
    writer.write_header([
        "ID",
        "Order Numbers",
        "Warehouse Code",
        "Invoice No",
        "Invoice Pattern",
        "Invoice Serial",
        "FPT Transaction ID",
        "Document Date",
        "Issue Date",
        "Status",
        "Buyer Name",
        "Buyer Tax Code",
        "Buyer Address",
        "Buyer Email",
        "Buyer Budget Relation ID",
        "Tax Authority Code",
        "Total Amount",
        "VAT Amount",
        "Grand Total",
        "Last Sync Attempt At",
        "Last Sync Error Message",
        "Secret Code",
        "Signing Date",
        "Created At",
        "Updated At",
    ])?;
    println!("Header written.\n");

    // Statistics
    let mut total_rows = 0;
    let mut batch_number = 0;
    let mut last_progress_update = Instant::now();
    let mut processing_times: Vec<Duration> = Vec::new();

    println!("Starting data export...\n");

    // Fetch and process in batches
    loop {
        let batch_start = Instant::now();

        // Fetch next batch using cursor
        let fetch_query = format!("FETCH {} FROM export_cursor", batch_size);
        let rows = transaction.query(&fetch_query, &[])?;

        if rows.is_empty() {
            println!("\nNo more data. Export complete.");
            break;
        }

        let batch_size_actual = rows.len();
        batch_number += 1;

        // Process batch
        for row in rows {
            // Extract e_invoice data with flexible types (handle NULLs and different DB schemas)
            let id: i64 = row.get(0);
            let order_numbers: Option<String> = row.try_get(1).ok().flatten();
            let warehouse_code: Option<String> = row.try_get(2).ok().flatten();
            let invoice_no: Option<String> = row.try_get(3).ok().flatten();
            let invoice_pattern: Option<String> = row.try_get(4).ok().flatten();
            let invoice_serial: Option<String> = row.try_get(5).ok().flatten();
            let fpt_transaction_id: Option<String> = row.try_get(6).ok().flatten();
            let document_date: Option<chrono::NaiveDate> = row.try_get(7).ok().flatten();
            let issue_date: Option<chrono::NaiveDateTime> = row.try_get(8).ok().flatten();
            let status: Option<String> = row.try_get(9).ok().flatten();
            let buyer_name: Option<String> = row.try_get(10).ok().flatten();
            let buyer_tax_code: Option<String> = row.try_get(11).ok().flatten();
            let buyer_address: Option<String> = row.try_get(12).ok().flatten();
            let buyer_email: Option<String> = row.try_get(13).ok().flatten();
            let buyer_budget_relation_id: Option<String> = row.try_get(14).ok().flatten();
            let tax_authority_code: Option<String> = row.try_get(15).ok().flatten();
            let total_amount: Option<f64> = row.try_get(16).ok().flatten();
            let vat_amount: Option<f64> = row.try_get(17).ok().flatten();
            let grand_total: Option<f64> = row.try_get(18).ok().flatten();
            let last_sync_attempt_at: Option<chrono::NaiveDateTime> =
                row.try_get(19).ok().flatten();
            let last_sync_error_message: Option<String> = row.try_get(20).ok().flatten();
            let secret_code: Option<String> = row.try_get(21).ok().flatten();
            let signing_date: Option<chrono::NaiveDateTime> = row.try_get(22).ok().flatten();
            let created_at: chrono::NaiveDateTime = row
                .try_get(23)
                .unwrap_or_else(|_| chrono::Utc::now().naive_utc());
            let updated_at: chrono::NaiveDateTime = row
                .try_get(24)
                .unwrap_or_else(|_| chrono::Utc::now().naive_utc());

            // MEMORY OPTIMIZATION: Format timestamps ONCE, avoid intermediate allocations
            let document_date_str = document_date
                .map(|d| d.format("%Y-%m-%d").to_string())
                .unwrap_or_default();
            let issue_date_str = issue_date
                .map(|d| d.format("%Y-%m-%d %H:%M:%S").to_string())
                .unwrap_or_default();
            let last_sync_attempt_str = last_sync_attempt_at
                .map(|d| d.format("%Y-%m-%d %H:%M:%S").to_string())
                .unwrap_or_default();
            let signing_date_str = signing_date
                .map(|d| d.format("%Y-%m-%d %H:%M:%S").to_string())
                .unwrap_or_default();
            let created_at_str = created_at.format("%Y-%m-%d %H:%M:%S").to_string();
            let updated_at_str = updated_at.format("%Y-%m-%d %H:%M:%S").to_string();

            // Write to Excel using typed values (40% faster than strings)
            writer.write_row_typed(&[
                CellValue::Int(id),
                CellValue::String(order_numbers.unwrap_or_default()),
                CellValue::String(warehouse_code.unwrap_or_default()),
                CellValue::String(invoice_no.unwrap_or_default()),
                CellValue::String(invoice_pattern.unwrap_or_default()),
                CellValue::String(invoice_serial.unwrap_or_default()),
                CellValue::String(fpt_transaction_id.unwrap_or_default()),
                CellValue::String(document_date_str),
                CellValue::String(issue_date_str),
                CellValue::String(status.unwrap_or_default()),
                CellValue::String(buyer_name.unwrap_or_default()),
                CellValue::String(buyer_tax_code.unwrap_or_default()),
                CellValue::String(buyer_address.unwrap_or_default()),
                CellValue::String(buyer_email.unwrap_or_default()),
                CellValue::String(buyer_budget_relation_id.unwrap_or_default()),
                CellValue::String(tax_authority_code.unwrap_or_default()),
                CellValue::Float(total_amount.unwrap_or(0.0)),
                CellValue::Float(vat_amount.unwrap_or(0.0)),
                CellValue::Float(grand_total.unwrap_or(0.0)),
                CellValue::String(last_sync_attempt_str),
                CellValue::String(last_sync_error_message.unwrap_or_default()),
                CellValue::String(secret_code.unwrap_or_default()),
                CellValue::String(signing_date_str),
                CellValue::String(created_at_str),
                CellValue::String(updated_at_str),
            ])?;
        }

        total_rows += batch_size_actual;

        let batch_duration = batch_start.elapsed();
        processing_times.push(batch_duration);

        // Keep only last 10 batch times for moving average
        if processing_times.len() > 10 {
            processing_times.remove(0);
        }

        // Progress update every 2 seconds
        if last_progress_update.elapsed() > Duration::from_secs(2) {
            let _avg_time_per_batch =
                processing_times.iter().sum::<Duration>() / processing_times.len() as u32;
            let rows_per_sec = (batch_size_actual as f64) / batch_duration.as_secs_f64();

            println!(
                "  Batch {:>4} | Rows: {:>10} | Speed: {:>6.0} rows/sec | Batch time: {:>6.2}s",
                batch_number,
                total_rows,
                rows_per_sec,
                batch_duration.as_secs_f64()
            );

            last_progress_update = Instant::now();
        }

        // Stop if we got less than requested (end of data)
        if batch_size_actual < batch_size as usize {
            break;
        }
    }

    // Close cursor and commit transaction
    transaction.execute("CLOSE export_cursor", &[])?;
    transaction.commit()?;

    // Save workbook
    println!("\nFinalizing Excel file...");
    writer.save()?;

    let total_duration = start.elapsed();

    // Final statistics
    println!("\n=== Export Statistics ===");
    println!("Total rows exported: {}", total_rows);
    println!("Total time: {:?}", total_duration);
    println!(
        "Average speed: {:.0} rows/sec",
        total_rows as f64 / total_duration.as_secs_f64()
    );
    println!("Number of batches: {}", batch_number);
    println!("Output file: {}", output_file);

    // Estimate file size
    if let Ok(metadata) = std::fs::metadata(output_file) {
        let size_mb = metadata.len() as f64 / 1_048_576.0;
        println!("File size: {:.2} MB", size_mb);
    }

    println!("\n✓ Export completed successfully!");

    Ok(())
}
