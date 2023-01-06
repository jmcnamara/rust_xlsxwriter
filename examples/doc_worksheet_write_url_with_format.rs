// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing a url with alternative format.

use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError, XlsxUnderline};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a format to use in the worksheet.
    let link_format = Format::new()
        .set_font_color(XlsxColor::Red)
        .set_underline(XlsxUnderline::Single);

    // Write a url with an alternative format.
    worksheet.write_url_with_format(0, 0, "https://www.rust-lang.org", &link_format)?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
