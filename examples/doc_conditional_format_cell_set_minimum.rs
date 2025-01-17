// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! Example of adding a cell type conditional formatting to a worksheet. Values
//! between 40 and 60 are highlighted in light green.

use rust_xlsxwriter::{
    ConditionalFormatCell, ConditionalFormatCellRule, Format, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some sample data.
    let data = [10, 80, 50, 10, 20, 60, 40, 70, 30, 40];

    worksheet.write_column(0, 0, data)?;

    // Add a format. Green fill with dark green text.
    let format = Format::new()
        .set_font_color("006100")
        .set_background_color("C6EFCE");

    // Write a conditional format over a range.
    let conditional_format = ConditionalFormatCell::new()
        .set_rule(ConditionalFormatCellRule::Between(40, 60))
        .set_format(format);

    worksheet.add_conditional_format(0, 0, 9, 0, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
