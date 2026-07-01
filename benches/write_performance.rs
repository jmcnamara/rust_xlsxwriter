// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! Benchmark tests for worksheet write performance.
//!
//! These tests measure xlsx file creation for several data combinations and
//! worksheet memory modes. The memory mode is selected at compile time so that
//! runs can be compared against a saved baseline:
//!
//! - default: `add_worksheet()`. The entire worksheet data structure is held in
//!   memory.
//! - `constant_memory` feature: `add_worksheet_with_constant_memory()`. The
//!   majority of the worksheet data is written to a temporary file and copied
//!   into the final xlsx/zip file.
//! - `constant_memory` feature + `BENCH_LOW_MEMORY=1`:
//!   `add_worksheet_with_low_memory()`: Like `constant_memory` mode but 1
//!   instance of each unique string is held in memory.
//!
//! To compare the default worksheet against constant/low memory modes:
//!
//! ```bash
//! # Default worksheet (baseline).
//! cargo bench --bench write_performance -- --save-baseline default
//!
//! # Constant memory mode, compared against the baseline.
//! cargo bench --bench write_performance --features constant_memory -- --baseline default
//!
//! # Low memory mode, compared against the baseline.
//! BENCH_LOW_MEMORY=1 cargo bench --bench write_performance --features constant_memory -- --baseline default
//! ```

use criterion::{black_box, criterion_group, criterion_main, BenchmarkId, Criterion, Throughput};
use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};

// All configurations use a fixed column count and vary the row count.
const COLS: u16 = 50;

// (rows, name) for 1K, 10K, 100K and 1M cells.
const CONFIGS: [(u32, &str); 4] = [
    (20, "1K cells"),
    (200, "10K cells"),
    (2_000, "100K cells"),
    (20_000, "1M cells"),
];

/// Benchmark number and string cells (50% of each).
fn bench_mixed(c: &mut Criterion) {
    run_group(c, "mixed", write_mixed);
}

/// Benchmark a number (12345.0) in every cell.
fn bench_numbers(c: &mut Criterion) {
    run_group(c, "numbers", write_numbers);
}

/// Benchmark a static string ("Foo") in every cell.
fn bench_string_static(c: &mut Criterion) {
    run_group(c, "string_static", write_string_static);
}

/// Benchmark a per-cell dynamic string ("Cell (row, col)") in every cell.
fn bench_string_dynamic(c: &mut Criterion) {
    run_group(c, "string_dynamic", write_string_dynamic);
}

/// Run a benchmark group over all cell-count configurations for a given data
/// combination.
fn run_group(
    c: &mut Criterion,
    name: &str,
    populate: fn(&mut Worksheet, u32, u16) -> Result<(), XlsxError>,
) {
    let mut group = c.benchmark_group(name);

    for (rows, size_name) in CONFIGS {
        let cell_count = rows * COLS as u32;
        group.throughput(Throughput::Elements(cell_count as u64));

        // Fewer samples for the larger, slower configs.
        let sample_size = if cell_count >= 1_000_000 {
            10
        } else if cell_count >= 100_000 {
            20
        } else {
            50
        };
        group.sample_size(sample_size);

        group.bench_with_input(BenchmarkId::from_parameter(size_name), &rows, |b, &rows| {
            b.iter(|| {
                write_workbook(black_box(rows), black_box(COLS), populate)
                    .expect("Failed to write workbook");
            });
        });
    }

    group.finish();
}

/// Create a workbook, populate it and save it to a temporary file.
fn write_workbook(
    rows: u32,
    cols: u16,
    populate: fn(&mut Worksheet, u32, u16) -> Result<(), XlsxError>,
) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    populate(add_benchmark_worksheet(&mut workbook), rows, cols)?;

    let temp_file =
        std::env::temp_dir().join(format!("bench_write_perf_{}.xlsx", std::process::id()));
    workbook.save(&temp_file)?;
    let _ = std::fs::remove_file(temp_file);

    Ok(())
}

/// Add a worksheet in the memory mode selected by feature flags / env var.
fn add_benchmark_worksheet(workbook: &mut Workbook) -> &mut Worksheet {
    #[cfg(not(feature = "constant_memory"))]
    return workbook.add_worksheet();

    #[cfg(feature = "constant_memory")]
    if std::env::var_os("BENCH_LOW_MEMORY").is_some() {
        workbook.add_worksheet_with_low_memory()
    } else {
        workbook.add_worksheet_with_constant_memory()
    }
}

/// Write numbers and strings, alternatively, to each cell in the range.
fn write_mixed(worksheet: &mut Worksheet, rows: u32, cols: u16) -> Result<(), XlsxError> {
    for row in 0..rows {
        for col in 0..cols {
            if col % 2 == 1 {
                worksheet.write_string(row, col, "Foo")?;
            } else {
                worksheet.write_number(row, col, 12345.0)?;
            }
        }
    }

    Ok(())
}

/// Write the same number to every cell.
fn write_numbers(worksheet: &mut Worksheet, rows: u32, cols: u16) -> Result<(), XlsxError> {
    for row in 0..rows {
        for col in 0..cols {
            worksheet.write_number(row, col, 12345.0)?;
        }
    }

    Ok(())
}

/// Write the same static string to every cell.
fn write_string_static(worksheet: &mut Worksheet, rows: u32, cols: u16) -> Result<(), XlsxError> {
    for row in 0..rows {
        for col in 0..cols {
            worksheet.write_string(row, col, "Foo")?;
        }
    }

    Ok(())
}

/// Write a dynamic string to every cell.
fn write_string_dynamic(worksheet: &mut Worksheet, rows: u32, cols: u16) -> Result<(), XlsxError> {
    for row in 0..rows {
        for col in 0..cols {
            worksheet.write_string(row, col, format!("Cell ({row}, {col})"))?;
        }
    }

    Ok(())
}

criterion_group!(
    benches,
    bench_numbers,
    bench_mixed,
    bench_string_static,
    bench_string_dynamic,
);
criterion_main!(benches);
