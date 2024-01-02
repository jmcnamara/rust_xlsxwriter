// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of adding a data bar type conditional formatting to a worksheet with
//! a user defined axis color.

use rust_xlsxwriter::{ConditionalFormatDataBar, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write the worksheet data.
    let data = [6, 4, 2, -2, -4, -6, -4, -2, 2, 4];
    worksheet.write_column(2, 1, data)?;
    worksheet.write_column(2, 3, data)?;

    // Write a standard Excel data bar.
    let conditional_format = ConditionalFormatDataBar::new();

    worksheet.add_conditional_format(2, 1, 11, 1, &conditional_format)?;

    // Write a data bar with a user defined axis color.
    let conditional_format = ConditionalFormatDataBar::new().set_axis_color("FF0000");

    worksheet.add_conditional_format(2, 3, 11, 3, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
