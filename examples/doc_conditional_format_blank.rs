// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of how to add a blank/non-blank conditional formatting to a
//! worksheet. Blank values are in light red. Non-blank values are in light
//! green. Note, that we invert the Blank rule to get Non-blank values.

use rust_xlsxwriter::{ConditionalFormatBlank, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some sample data.
    let data = [
        "Not blank",
        "",
        "",
        "Not blank",
        "Not blank",
        "",
        "Not blank",
        "Not blank",
        "",
        "Not blank",
        "",
        "Not blank",
    ];
    worksheet.write_column(0, 0, data)?;

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
    let conditional_format = ConditionalFormatBlank::new().set_format(format1);

    worksheet.add_conditional_format(0, 0, 11, 0, &conditional_format)?;

    // Invert the blank conditional format to show non-blank values.
    let conditional_format = ConditionalFormatBlank::new().invert().set_format(format2);

    worksheet.add_conditional_format(0, 0, 11, 0, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
