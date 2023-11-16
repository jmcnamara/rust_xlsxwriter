// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! Example of adding data bar type conditional formatting to a worksheet.

use rust_xlsxwriter::{
    ConditionalFormatDataBar, ConditionalFormatDataBarDirection, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write some captions.
    worksheet.write(1, 1, "Default")?;
    worksheet.write(1, 3, "Default negative")?;
    worksheet.write(1, 5, "User color")?;
    worksheet.write(1, 7, "Changed direction")?;

    // Write the worksheet data.
    let data1 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    let data2 = [6, 4, 2, -2, -4, -6, -4, -2, 2, 4];
    worksheet.write_column(2, 1, data1)?;
    worksheet.write_column(2, 3, data2)?;
    worksheet.write_column(2, 5, data1)?;
    worksheet.write_column(2, 7, data1)?;

    // Write a standard Excel data bar.
    let conditional_format = ConditionalFormatDataBar::new();

    worksheet.add_conditional_format(2, 1, 11, 1, &conditional_format)?;

    // Write a standard Excel data bar with negative data
    let conditional_format = ConditionalFormatDataBar::new();

    worksheet.add_conditional_format(2, 3, 11, 3, &conditional_format)?;

    // Write a data bar with a user defined fill color.
    let conditional_format = ConditionalFormatDataBar::new().set_fill_color("009933");

    worksheet.add_conditional_format(2, 5, 11, 5, &conditional_format)?;

    // Write a data bar with the direction changed.
    let conditional_format = ConditionalFormatDataBar::new()
        .set_direction(ConditionalFormatDataBarDirection::RightToLeft);

    worksheet.add_conditional_format(2, 7, 11, 7, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
