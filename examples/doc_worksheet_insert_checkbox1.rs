// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! This example demonstrates adding adding checkbox boolean values to a
//! worksheet.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Insert some boolean checkboxes to the worksheet.
    worksheet.insert_checkbox(2, 2, false)?;
    worksheet.insert_checkbox(3, 2, true)?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
