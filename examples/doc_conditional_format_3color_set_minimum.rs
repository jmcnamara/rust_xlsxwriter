// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! Example of adding 3 color scale type conditional formatting to a worksheet
//! with user defined minimum and maximum values.

use rust_xlsxwriter::{ConditionalFormat3ColorScale, ConditionalFormatType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write the worksheet data.
    let data = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    worksheet.write_column(2, 1, data)?;
    worksheet.write_column(2, 3, data)?;

    // Write a 3 color scale formats with standard Excel colors. The conditional
    // format is applied from the lowest to the highest value.
    let conditional_format = ConditionalFormat3ColorScale::new();

    worksheet.add_conditional_format(2, 1, 11, 1, &conditional_format)?;

    // Write a 3 color scale formats with standard Excel colors but a user
    // defined range. Values <= 3 will be shown with the minimum color while
    // values >= 7 will be shown with the maximum color.
    let conditional_format = ConditionalFormat3ColorScale::new()
        .set_minimum(ConditionalFormatType::Number, 3)
        .set_maximum(ConditionalFormatType::Number, 7);

    worksheet.add_conditional_format(2, 3, 11, 3, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
