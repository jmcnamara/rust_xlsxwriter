// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! Example of adding a data bar type conditional formatting to a worksheet
//! with different axis positions.

use rust_xlsxwriter::{
    ConditionalFormatDataBar, ConditionalFormatDataBarAxisPosition, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write the worksheet data.
    let data1 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    let data2 = [6, 4, 2, -2, -4, -6, -4, -2, 2, 4];
    worksheet.write_column(2, 1, data1)?;
    worksheet.write_column(2, 3, data1)?;
    worksheet.write_column(2, 5, data2)?;
    worksheet.write_column(2, 7, data2)?;

    // Write a standard Excel data bar.
    let conditional_format = ConditionalFormatDataBar::new();

    worksheet.add_conditional_format(2, 1, 11, 1, &conditional_format)?;

    // Write a data bar with a midpoint axis.
    let conditional_format = ConditionalFormatDataBar::new()
        .set_axis_position(ConditionalFormatDataBarAxisPosition::Midpoint);

    worksheet.add_conditional_format(2, 3, 11, 3, &conditional_format)?;

    // Write a standard Excel data bar with negative data
    let conditional_format = ConditionalFormatDataBar::new();

    worksheet.add_conditional_format(2, 5, 11, 5, &conditional_format)?;

    // Write a data bar without an axis.
    let conditional_format = ConditionalFormatDataBar::new()
        .set_axis_position(ConditionalFormatDataBarAxisPosition::None);

    worksheet.add_conditional_format(2, 7, 11, 7, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
