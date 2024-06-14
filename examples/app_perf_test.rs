// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Simple performance test for rust_xlsxwriter.

use rust_xlsxwriter::{Workbook, XlsxError};
use std::env;

fn main() -> Result<(), XlsxError> {
    let args: Vec<String> = env::args().collect();

    let col_max = 50;
    let row_max = match args.get(1) {
        Some(arg) => arg.parse::<u32>().unwrap_or(4_000),
        None => 4_000,
    };

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

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

    Ok(())
}
