// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! Example of how to add conditional formatting to a worksheet using the
//! rust_xlsxwriter library.
//!
//! Conditional formatting allows you to apply a format to a cell or a range of
//! cells based on certain criteria.

use rust_xlsxwriter::{
    ConditionalFormatCell, ConditionalFormatCellCriteria, Format, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a format. Light red fill with dark red text.
    let format1 = Format::new()
        .set_font_color("9C0006")
        .set_background_color("FFC7CE");

    // Add a format. Green fill with dark green text.
    let format2 = Format::new()
        .set_font_color("006100")
        .set_background_color("C6EFCE");

    // some sample data to run the conditional formatting against.
    let data = [
        [90, 80, 50, 10, 20, 90, 40, 90, 30, 40],
        [20, 10, 90, 100, 30, 60, 70, 60, 50, 90],
        [10, 50, 60, 50, 20, 50, 80, 30, 40, 60],
        [10, 90, 20, 40, 10, 40, 50, 70, 90, 50],
        [70, 100, 10, 90, 10, 10, 20, 100, 100, 40],
        [20, 60, 10, 100, 30, 10, 20, 60, 100, 10],
        [10, 60, 10, 80, 100, 80, 30, 30, 70, 40],
        [30, 90, 60, 10, 10, 100, 40, 40, 30, 40],
        [80, 90, 10, 20, 20, 50, 80, 20, 60, 90],
        [60, 80, 30, 30, 10, 50, 80, 60, 50, 30],
    ];

    // -----------------------------------------------------------------------
    // Example 1. Cell conditional formatting.
    // -----------------------------------------------------------------------
    let caption = "Cells with values >= 50 are in light red. Values < 50 are in light green.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    worksheet.write_row_matrix(2, 1, data)?;

    // Set the column widths for clarity.
    for col_num in 1..=10u16 {
        worksheet.set_column_width(col_num, 6)?;
    }

    // Write a conditional format over a range.
    let conditional_format = ConditionalFormatCell::new()
        .set_criteria(ConditionalFormatCellCriteria::GreaterThanOrEqualTo)
        .set_value(50)
        .set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Write another conditional format over the same range.
    let conditional_format = ConditionalFormatCell::new()
        .set_criteria(ConditionalFormatCellCriteria::LessThan)
        .set_value(50)
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Example 2. Cell conditional formatting with between ranges.
    // -----------------------------------------------------------------------
    let caption =
        "Values between 30 and 70 are in light red. Values outside that range are in light green.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    worksheet.write_row_matrix(2, 1, data)?;

    // Set the column widths for clarity.
    for col_num in 1..=10u16 {
        worksheet.set_column_width(col_num, 6)?;
    }

    // Write a conditional format over a range.
    let conditional_format = ConditionalFormatCell::new()
        .set_criteria(ConditionalFormatCellCriteria::Between)
        .set_minimum(30)
        .set_maximum(70)
        .set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Write another conditional format over the same range.
    let conditional_format = ConditionalFormatCell::new()
        .set_criteria(ConditionalFormatCellCriteria::NotBetween)
        .set_minimum(30)
        .set_maximum(70)
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Save and close the file.
    // -----------------------------------------------------------------------
    workbook.save("conditional_formats.xlsx")?;

    Ok(())
}
