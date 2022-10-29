// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting a number format that appears
//! differently in different locales.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    // Add a format.
    let currency_format = Format::new().set_num_format("#,##0.00");

    worksheet.write_number(0, 0, 1234.56, &currency_format)?;

    workbook.save("number_format.xlsx")?;

    Ok(())
}
