// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! Example of adding a data bar type conditional formatting to a worksheet with
//! user defined minimum and maximum values.

use rust_xlsxwriter::{ConditionalFormatDataBar, ConditionalFormatType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write the worksheet data.
    let data = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    worksheet.write_column(2, 1, data)?;
    worksheet.write_column(2, 3, data)?;

    // Write a standard Excel data bar. The conditional format is applied over
    // the full range of values from minimum to maximum.
    let conditional_format = ConditionalFormatDataBar::new();

    worksheet.add_conditional_format(2, 1, 11, 1, &conditional_format)?;

    // Write a data bar a user defined range. Values <= 3 are shown with zero
    // bar width while values >= 7 are shown with the maximum bar width.
    let conditional_format = ConditionalFormatDataBar::new()
        .set_minimum(ConditionalFormatType::Number, 3)
        .set_maximum(ConditionalFormatType::Number, 7);

    worksheet.add_conditional_format(2, 3, 11, 3, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
