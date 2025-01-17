// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates auto-fitting the worksheet column widths
//! based on the data in the columns.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add some data
    worksheet.write_string(0, 0, "Hello")?;
    worksheet.write_string(0, 1, "Hello")?;
    worksheet.write_string(1, 1, "Hello World")?;
    worksheet.write_number(0, 2, 123)?;
    worksheet.write_number(0, 3, 123456)?;

    // Autofit the columns.
    worksheet.autofit();

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
