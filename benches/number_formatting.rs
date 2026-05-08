// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! Criterion benchmark for comparing number formatting performance
//! with default, `ryu`, and `zmij` features.
//!
//! To run benchmarks with different features:
//!
//! ```bash
//! # Default (no feature)
//! cargo bench --bench number_formatting
//!
//! # With ryu feature
//! cargo bench --bench number_formatting --features ryu
//!
//! # With zmij feature
//! cargo bench --bench number_formatting --features zmij
//! ```

use criterion::{black_box, criterion_group, criterion_main, BenchmarkId, Criterion, Throughput};
use rust_xlsxwriter::{Workbook, XlsxError};
use std::io::Cursor;

/// Benchmark writing numbers to a worksheet
fn bench_write_numbers(c: &mut Criterion) {
    let mut group = c.benchmark_group("write_numbers");

    // Test with different cell counts
    let configs: Vec<(u32, u16, &str)> = vec![
        (50, 100, "5K cells"),      // 5,000 cells
        (100, 200, "20K cells"),    // 20,000 cells
        (200, 250, "50K cells"),    // 50,000 cells
        (500, 400, "200K cells"),   // 200,000 cells
        (1000, 500, "500K cells"),  // 500,000 cells
    ];

    for (rows, cols, name) in configs {
        let cell_count = rows * cols as u32;
        group.throughput(Throughput::Elements(cell_count as u64));

        group.bench_with_input(BenchmarkId::from_parameter(name), &(rows, cols), |b, &(rows, cols)| {
            b.iter(|| {
                write_numbers_to_workbook(black_box(rows), black_box(cols))
                    .expect("Failed to write workbook");
            });
        });
    }

    group.finish();
}

/// Benchmark writing numbers to memory (faster, focuses on formatting)
fn bench_write_numbers_to_memory(c: &mut Criterion) {
    let mut group = c.benchmark_group("write_numbers_to_memory");

    // Test with different cell counts
    let configs: Vec<(u32, u16, &str)> = vec![
        (50, 100, "5K cells"),      // 5,000 cells
        (100, 200, "20K cells"),    // 20,000 cells
        (200, 250, "50K cells"),    // 50,000 cells
        (500, 400, "200K cells"),   // 200,000 cells
    ];

    for (rows, cols, name) in configs {
        let cell_count = rows * cols as u32;
        group.throughput(Throughput::Elements(cell_count as u64));

        group.bench_with_input(BenchmarkId::from_parameter(name), &(rows, cols), |b, &(rows, cols)| {
            b.iter(|| {
                write_numbers_to_memory(black_box(rows), black_box(cols))
                    .expect("Failed to write workbook");
            });
        });
    }

    group.finish();
}

/// Benchmark mixed data types (numbers and strings)
fn bench_write_mixed_data(c: &mut Criterion) {
    let mut group = c.benchmark_group("write_mixed_data");

    let configs: Vec<(u32, u16, &str)> = vec![
        (50, 100, "5K cells"),      // 5,000 cells
        (200, 250, "50K cells"),    // 50,000 cells
    ];

    for (rows, cols, name) in configs {
        let cell_count = rows * cols as u32;
        group.throughput(Throughput::Elements(cell_count as u64));

        group.bench_with_input(BenchmarkId::from_parameter(name), &(rows, cols), |b, &(rows, cols)| {
            b.iter(|| {
                write_mixed_data_to_workbook(black_box(rows), black_box(cols))
                    .expect("Failed to write workbook");
            });
        });
    }

    group.finish();
}

/// Helper function to write numbers to a workbook and save to a temporary file
fn write_numbers_to_workbook(rows: u32, cols: u16) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    for row in 0..rows {
        for col in 0..cols {
            worksheet.write_number(row, col, 12345.67890)?;
        }
    }

    // Use a temporary file to avoid disk I/O affecting the benchmark
    let temp_dir = std::env::temp_dir();
    let temp_file = temp_dir.join(format!("bench_test_{}.xlsx", std::process::id()));
    workbook.save(&temp_file)?;
    let _ = std::fs::remove_file(temp_file);

    Ok(())
}

/// Helper function to write numbers to memory (buffer)
fn write_numbers_to_memory(rows: u32, cols: u16) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    for row in 0..rows {
        for col in 0..cols {
            worksheet.write_number(row, col, 12345.67890)?;
        }
    }

    // Save to memory buffer
    let mut buffer = Cursor::new(Vec::new());
    workbook.save_to_writer(&mut buffer)?;

    Ok(())
}

/// Helper function to write mixed data (alternating strings and numbers)
fn write_mixed_data_to_workbook(rows: u32, cols: u16) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    for row in 0..rows {
        for col in 0..cols {
            if col % 2 == 0 {
                worksheet.write_number(row, col, 12345.67890)?;
            } else {
                worksheet.write_string(row, col, "Sample text")?;
            }
        }
    }

    // Save to memory buffer
    let mut buffer = Cursor::new(Vec::new());
    workbook.save_to_writer(&mut buffer)?;

    Ok(())
}

criterion_group!(benches, bench_write_numbers, bench_write_numbers_to_memory, bench_write_mixed_data);
criterion_main!(benches);
