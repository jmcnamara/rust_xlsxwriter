// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing a URL with alternative format.

use rust_xlsxwriter::{Color, Format, FormatUnderline, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a format to use in the worksheet.
    let link_format = Format::new()
        .set_font_color(Color::Red)
        .set_underline(FormatUnderline::Single);

    // Write a URL with an alternative format.
    worksheet.write_url_with_format(0, 0, "https://www.rust-lang.org", &link_format)?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
