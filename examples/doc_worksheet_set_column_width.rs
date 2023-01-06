// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the width of columns in Excel.
use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add some text.
    worksheet.write_string_only(0, 0, "Normal")?;
    worksheet.write_string_only(0, 2, "Wider")?;
    worksheet.write_string_only(0, 4, "Narrower")?;

    // Set the column width in Excel character units.
    worksheet.set_column_width(2, 16)?;
    worksheet.set_column_width(4, 4)?;
    worksheet.set_column_width(5, 4)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
