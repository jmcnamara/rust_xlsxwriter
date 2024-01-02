// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of adding a Formula type conditional formatting to a worksheet. This
//! example demonstrate the effect of changing the absolute/relative anchor in
//! the target cell.

use rust_xlsxwriter::{ConditionalFormatFormula, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a format. Green fill with dark green text.
    let format = Format::new()
        .set_font_color("006100")
        .set_background_color("C6EFCE");

    // Add some sample data.
    let data = [
        [34, 73, 39, 32, 75, 48, 75, 66],
        [5, 24, 1, 84, 54, 62, 60, 3],
        [28, 79, 97, 13, 85, 93, 93, 22],
        [27, 71, 40, 17, 18, 79, 90, 93],
        [88, 25, 33, 23, 67, 1, 59, 79],
        [23, 100, 20, 88, 29, 33, 38, 54],
        [7, 57, 88, 28, 10, 26, 37, 7],
        [53, 78, 1, 96, 26, 45, 47, 33],
        [60, 54, 81, 66, 81, 90, 80, 93],
        [70, 5, 46, 14, 71, 19, 66, 36],
    ];

    // Add a new worksheet and write the sample data.
    let worksheet = workbook.add_worksheet();
    worksheet.write_row_matrix(2, 1, data)?;

    // The rule is applied to each cell in the range.
    let conditional_format = ConditionalFormatFormula::new()
        .set_rule("=ISEVEN(B3)")
        .set_format(&format);

    worksheet.add_conditional_format(2, 1, 11, 8, &conditional_format)?;

    // Add a new worksheet and write the sample data.
    let worksheet = workbook.add_worksheet();
    worksheet.write_row_matrix(2, 1, data)?;

    // The rule is applied to each row based on the first row in the column.
    let conditional_format = ConditionalFormatFormula::new()
        .set_rule("=ISEVEN($B3)")
        .set_format(&format);

    worksheet.add_conditional_format(2, 1, 11, 8, &conditional_format)?;

    // Add a new worksheet and write the sample data.
    let worksheet = workbook.add_worksheet();
    worksheet.write_row_matrix(2, 1, data)?;

    // The rule is applied to each column based on the first cell in the column.
    let conditional_format = ConditionalFormatFormula::new()
        .set_rule("=ISEVEN(B$3)")
        .set_format(&format);

    worksheet.add_conditional_format(2, 1, 11, 8, &conditional_format)?;

    // Add a new worksheet and write the sample data.
    let worksheet = workbook.add_worksheet();
    worksheet.write_row_matrix(2, 1, data)?;

    // The rule is applied to the entire range based on the first cell in the range.
    let conditional_format = ConditionalFormatFormula::new()
        .set_rule("=ISEVEN($B$3)")
        .set_format(format);

    worksheet.add_conditional_format(2, 1, 11, 8, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
