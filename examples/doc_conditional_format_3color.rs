// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of adding 3 color scale type conditional formatting to a worksheet.
//! Note, the colors in the first example (red to yellow to green) are the
//! default colors and could be omitted.

use rust_xlsxwriter::{ConditionalFormat3ColorScale, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write the worksheet data.
    let data = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    worksheet.write_column(2, 1, data)?;
    worksheet.write_column(2, 3, data)?;
    worksheet.write_column(2, 5, data)?;
    worksheet.write_column(2, 7, data)?;
    worksheet.write_column(2, 9, data)?;
    worksheet.write_column(2, 11, data)?;

    // Set the column widths for clarity.
    worksheet.set_column_range_width(0, 12, 6)?;

    // Write 3 color scale formats with standard Excel colors.
    let conditional_format = ConditionalFormat3ColorScale::new()
        .set_minimum_color("F8696B")
        .set_midpoint_color("FFEB84")
        .set_maximum_color("63BE7B");

    worksheet.add_conditional_format(2, 1, 11, 1, &conditional_format)?;

    let conditional_format = ConditionalFormat3ColorScale::new()
        .set_minimum_color("63BE7B")
        .set_midpoint_color("FFEB84")
        .set_maximum_color("F8696B");

    worksheet.add_conditional_format(2, 3, 11, 3, &conditional_format)?;

    let conditional_format = ConditionalFormat3ColorScale::new()
        .set_minimum_color("F8696B")
        .set_midpoint_color("FCFCFF")
        .set_maximum_color("63BE7B");

    worksheet.add_conditional_format(2, 5, 11, 5, &conditional_format)?;

    let conditional_format = ConditionalFormat3ColorScale::new()
        .set_minimum_color("63BE7B")
        .set_midpoint_color("FCFCFF")
        .set_maximum_color("F8696B");

    worksheet.add_conditional_format(2, 7, 11, 7, &conditional_format)?;

    let conditional_format = ConditionalFormat3ColorScale::new()
        .set_minimum_color("F8696B")
        .set_midpoint_color("FCFCFF")
        .set_maximum_color("5A8AC6");

    worksheet.add_conditional_format(2, 9, 11, 9, &conditional_format)?;

    let conditional_format = ConditionalFormat3ColorScale::new()
        .set_minimum_color("5A8AC6")
        .set_midpoint_color("FCFCFF")
        .set_maximum_color("F8696B");

    worksheet.add_conditional_format(2, 11, 11, 11, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
