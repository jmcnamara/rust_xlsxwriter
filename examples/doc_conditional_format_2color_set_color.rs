// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! Example of adding a 2 color scale type conditional formatting to a worksheet
//! with user defined minimum and maximum colors.

use rust_xlsxwriter::{ConditionalFormat2ColorScale, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write the worksheet data.
    let data = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    worksheet.write(1, 1, "Default colors")?;
    worksheet.write(1, 3, "User colors")?;
    worksheet.write_column(2, 1, data)?;
    worksheet.write_column(2, 3, data)?;

    // Write a 2 color scale formats with standard Excel colors.
    let conditional_format = ConditionalFormat2ColorScale::new();

    worksheet.add_conditional_format(2, 1, 11, 1, &conditional_format)?;

    // Write a 2 color scale formats with user defined colors. This reverses the
    // default colors.
    let conditional_format = ConditionalFormat2ColorScale::new()
        .set_minimum_color("63BE7B")
        .set_maximum_color("FFEF9C");

    worksheet.add_conditional_format(2, 3, 11, 3, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
