// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! Example of setting the default theme for a workbook to a user supplied
//! custom theme using the `rust_xlsxwriter` library. The theme xml file is
//! extracted from an Excel xlsx file.

use rust_xlsxwriter::{FontScheme, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a custom theme to the workbook.
    workbook.use_custom_theme("tests/input/themes/technic.xml")?;

    // Create a new default format to match the custom theme. Note, that the
    // scheme is set to "Body" to indicate that the font is part of the theme.
    let format = Format::new()
        .set_font_name("Arial")
        .set_font_size(11)
        .set_font_scheme(FontScheme::Body);

    // Add the default format for the workbook.
    workbook.set_default_format(&format, 19, 72)?;

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some text to demonstrate the changed theme.
    worksheet.write(0, 0, "Hello")?;

    // Save the workbook to disk.
    workbook.save("theme_custom.xlsx")?;

    Ok(())
}
