// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! Example of adding 3 color scale type conditional formatting to a worksheet
//! with user defined minimum, midpoint and maximum colors.

use rust_xlsxwriter::{ConditionalFormat3ColorScale, Workbook, XlsxError};

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

    // Write a 3 color scale formats with standard Excel colors.
    let conditional_format = ConditionalFormat3ColorScale::new();

    worksheet.add_conditional_format(2, 1, 11, 1, &conditional_format)?;

    // Write a 3 color scale formats with user defined colors. This reverses the
    // default colors.
    let conditional_format = ConditionalFormat3ColorScale::new()
        .set_minimum_color("63BE7B")
        .set_midpoint_color("FFEB84")
        .set_maximum_color("F8696B");

    worksheet.add_conditional_format(2, 3, 11, 3, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
