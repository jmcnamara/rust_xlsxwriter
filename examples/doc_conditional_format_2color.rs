// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! Example of adding 2 color scale type conditional formatting to a worksheet.
//! Note, the colors in the first example could be omitted since they are the
//! default colors.

use rust_xlsxwriter::{ConditionalFormat2ColorScale, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write the worksheet data.
    let scale_data = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    worksheet.write_column(2, 1, scale_data)?;
    worksheet.write_column(2, 3, scale_data)?;
    worksheet.write_column(2, 5, scale_data)?;
    worksheet.write_column(2, 7, scale_data)?;

    // Write 2 color scale formats with standard Excel colors.
    let conditional_format = ConditionalFormat2ColorScale::new()
        .set_minimum_color("FCFCFF")
        .set_maximum_color("63BE7B");

    worksheet.add_conditional_format(2, 1, 11, 1, &conditional_format)?;

    let conditional_format = ConditionalFormat2ColorScale::new()
        .set_minimum_color("63BE7B")
        .set_maximum_color("FCFCFF");

    worksheet.add_conditional_format(2, 3, 11, 3, &conditional_format)?;

    let conditional_format = ConditionalFormat2ColorScale::new()
        .set_minimum_color("FFEF9C")
        .set_maximum_color("63BE7B");

    worksheet.add_conditional_format(2, 5, 11, 5, &conditional_format)?;

    let conditional_format = ConditionalFormat2ColorScale::new()
        .set_minimum_color("63BE7B")
        .set_maximum_color("FFEF9C");

    worksheet.add_conditional_format(2, 7, 11, 7, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
