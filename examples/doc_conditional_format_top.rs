// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of how to add Top and Bottom conditional formatting to a worksheet.
//! Top 10 values are in light red. Bottom 10 values are in light green.

use rust_xlsxwriter::{
    ConditionalFormatTop, ConditionalFormatTopRule, Format, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some sample data.
    let data = [
        [34, 72, 38, 30, 75, 48, 75, 66, 84, 86],
        [6, 24, 1, 84, 54, 62, 60, 3, 26, 59],
        [28, 79, 97, 13, 85, 93, 93, 22, 5, 14],
        [27, 71, 40, 17, 18, 79, 90, 93, 29, 47],
        [88, 25, 33, 23, 67, 1, 59, 79, 47, 36],
        [24, 100, 20, 88, 29, 33, 38, 54, 54, 88],
        [6, 57, 88, 28, 10, 26, 37, 7, 41, 48],
        [52, 78, 1, 96, 26, 45, 47, 33, 96, 36],
        [60, 54, 81, 66, 81, 90, 80, 93, 12, 55],
        [70, 5, 46, 14, 71, 19, 66, 36, 41, 21],
    ];
    worksheet.write_row_matrix(2, 1, data)?;

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
    let conditional_format = ConditionalFormatTop::new()
        .set_rule(ConditionalFormatTopRule::Top(10))
        .set_format(format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Invert the Top conditional format to show Bottom values.
    let conditional_format = ConditionalFormatTop::new()
        .set_rule(ConditionalFormatTopRule::Bottom(10))
        .set_format(format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
