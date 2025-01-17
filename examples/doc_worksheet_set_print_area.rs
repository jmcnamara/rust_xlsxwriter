// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the print area for several worksheets.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet1 = workbook.add_worksheet();
    // Set the print area to "A1:M32"
    worksheet1.set_print_area(0, 0, 31, 12)?;

    let worksheet2 = workbook.add_worksheet();
    // Set the print area to "1:32"
    worksheet2.set_print_area(0, 0, 31, 16_383)?;

    let worksheet3 = workbook.add_worksheet();
    // Set the print area to "A:M"
    worksheet3.set_print_area(0, 0, 1_048_575, 12)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
