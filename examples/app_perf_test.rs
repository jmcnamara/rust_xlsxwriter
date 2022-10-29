// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! Simple performance test for rust_xlsxwriter.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let row_max = 4000;
    let col_max = 50;

    for row in 0..row_max {
        for col in 0..col_max {
            if col % 2 == 1 {
                worksheet.write_string_only(row, col, "Foo")?;
            } else {
                worksheet.write_number_only(row, col, 12345.0)?;
            }
        }
    }
    workbook.save("rust_perf_test.xlsx")?;

    Ok(())
}
