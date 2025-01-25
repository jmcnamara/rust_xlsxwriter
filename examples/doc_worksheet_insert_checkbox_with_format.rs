// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! This example demonstrates adding adding a checkbox boolean value to a
//! worksheet along with a cell format.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create some cell formats with different colors.
    let format1 = Format::new().set_background_color("FFC7CE");
    let format2 = Format::new().set_background_color("C6EFCE");

    // Insert some boolean checkboxes to the worksheet.
    worksheet.insert_checkbox_with_format(2, 2, false, &format1)?;
    worksheet.insert_checkbox_with_format(3, 2, true, &format2)?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
