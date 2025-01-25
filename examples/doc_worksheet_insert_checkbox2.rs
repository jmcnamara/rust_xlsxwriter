// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! This example demonstrates adding adding checkbox boolean values to a
//! worksheet by making use of the Excel feature that a checkbox is actually a
//! boolean value with a special format.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a checkbox format.
    let format = Format::new().set_checkbox();

    // Insert some boolean checkboxes to the worksheet.
    worksheet.write_boolean_with_format(2, 2, false, &format)?;
    worksheet.write_boolean_with_format(3, 2, true, &format)?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
