// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting a currency format for a worksheet
//! cell.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    // Add a format.
    let currency_format = Format::new().set_num_format("[$$-409]#,##0.00");

    worksheet.write_number_with_format(0, 0, 1234.56, &currency_format)?;

    workbook.save("currency_format.xlsx")?;

    Ok(())
}
