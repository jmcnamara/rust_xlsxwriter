// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of how to add a error/non-error conditional formatting to a
//! worksheet. Error values are in light red. Non-error values are in light
//! green. Note, that we invert the Error rule to get Non-error values.

use rust_xlsxwriter::{ConditionalFormatError, Format, Formula, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some sample data.
    worksheet.write(0, 0, Formula::new("=1/1"))?;
    worksheet.write(1, 0, Formula::new("=1/0"))?;
    worksheet.write(2, 0, Formula::new("=1/0"))?;
    worksheet.write(3, 0, Formula::new("=1/1"))?;
    worksheet.write(4, 0, Formula::new("=1/1"))?;
    worksheet.write(5, 0, Formula::new("=1/0"))?;
    worksheet.write(6, 0, Formula::new("=1/1"))?;
    worksheet.write(7, 0, Formula::new("=1/1"))?;
    worksheet.write(8, 0, Formula::new("=1/0"))?;
    worksheet.write(9, 0, Formula::new("=1/1"))?;
    worksheet.write(10, 0, Formula::new("=1/0"))?;
    worksheet.write(11, 0, Formula::new("=1/1"))?;

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 10, 6)?;

    // Add a format. Light red fill with dark red text.
    let format1 = Format::new()
        .set_font_color("9C0006")
        .set_background_color("FFC7CE");

    // Add a format. Green fill with dark green text.
    let format2 = Format::new()
        .set_font_color("006100")
        .set_background_color("C6EFCE");

    // Write a conditional format over a range.
    let conditional_format = ConditionalFormatError::new().set_format(format1);

    worksheet.add_conditional_format(0, 0, 11, 0, &conditional_format)?;

    // Invert the error conditional format to show non-error values.
    let conditional_format = ConditionalFormatError::new().invert().set_format(format2);

    worksheet.add_conditional_format(0, 0, 11, 0, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
