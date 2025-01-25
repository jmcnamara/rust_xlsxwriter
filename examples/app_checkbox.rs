// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of adding checkbox boolean values to a worksheet using the
//! rust_xlsxwriter library.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create some formats to use in the worksheet.
    let bold = Format::new().set_bold();
    let light_red = Format::new().set_background_color("FFC7CE");
    let light_green = Format::new().set_background_color("C6EFCE");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set the column width for clarity.
    worksheet.set_column_width(0, 30)?;

    // Write some descriptions.
    worksheet.write_with_format(1, 0, "Some simple checkboxes:", &bold)?;
    worksheet.write_with_format(4, 0, "Some checkboxes with cell formats:", &bold)?;

    // Insert some boolean checkboxes to the worksheet.
    worksheet.insert_checkbox(1, 1, false)?;
    worksheet.insert_checkbox(2, 1, true)?;

    // Insert some checkboxes with cell formats.
    worksheet.insert_checkbox_with_format(4, 1, false, &light_red)?;
    worksheet.insert_checkbox_with_format(5, 1, true, &light_green)?;

    // Save the file to disk.
    workbook.save("checkbox.xlsx")?;

    Ok(())
}
