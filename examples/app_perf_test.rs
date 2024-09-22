// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Simple performance test program for rust_xlsxwriter.
//!
//! It writes alternate cells of strings and numbers.
//! It defaults to 4,000 rows x 40 columns.
//!
//! The number of rows and the "constant memory" mode can be optionally set.
//!
//! usage: ./target/release/examples/app_perf_test [num_rows]
//! [--constant-memory]
//!

use rust_xlsxwriter::{Workbook, XlsxError};
use std::{env, time::Instant};

fn main() -> Result<(), XlsxError> {
    let args: Vec<String> = env::args().collect();

    // Set some size arguments, optionally from the command line.
    let col_max = 50;
    let row_max = match args.get(1) {
        Some(arg) => arg.parse::<u32>().unwrap_or(4_000),
        None => 4_000,
    };
    let constant_memory = args.get(2).is_some();

    let start_time = Instant::now();

    // Create the workbook and fill in the required cell data.
    let mut workbook = Workbook::new();
    let worksheet = if constant_memory {
        workbook.add_worksheet_with_constant_memory()
    } else {
        workbook.add_worksheet()
    };

    for row in 0..row_max {
        for col in 0..col_max {
            if col % 2 == 1 {
                worksheet.write_string(row, col, "Foo")?;
            } else {
                worksheet.write_number(row, col, 12345.0)?;
            }
        }
    }
    workbook.save("rust_perf_test.xlsx")?;

    // Calculate and print the metrics.
    let time = (start_time.elapsed().as_millis() as f64) / 1000.0;

    println!("Wrote:  {row_max} rows x {col_max} cols. Constant memory = {constant_memory}.");
    println!("Time:   {time:.3} seconds.");

    Ok(())
}
