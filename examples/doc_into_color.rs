// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of the different types of color syntax that is supported by the
//! [`Into`] [`Color`] trait.

use rust_xlsxwriter::{Color, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    // Widen the column for clarity.
    worksheet.set_column_width_pixels(0, 80)?;

    // Some examples with named color enum values.
    let color_format = Format::new().set_background_color(Color::Green);
    worksheet.write_string(0, 0, "Green")?;
    worksheet.write_blank(0, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Red);
    worksheet.write_string(1, 0, "Red")?;
    worksheet.write_blank(1, 1, &color_format)?;

    // Write a RGB color using the Color::RGB() enum method.
    let color_format = Format::new().set_background_color(Color::RGB(0xFF7F50));
    worksheet.write_string(2, 0, "#FF7F50")?;
    worksheet.write_blank(2, 1, &color_format)?;

    // Write a RGB color with the shorter Html string variant.
    let color_format = Format::new().set_background_color("#6495ED");
    worksheet.write_string(3, 0, "#6495ED")?;
    worksheet.write_blank(3, 1, &color_format)?;

    // Write a RGB color with a Html string (but without the `#`).
    let color_format = Format::new().set_background_color("DCDCDC");
    worksheet.write_string(4, 0, "#DCDCDC")?;
    worksheet.write_blank(4, 1, &color_format)?;

    // Write a RGB color with the optional u32 variant.
    let color_format = Format::new().set_background_color(0xDAA520);
    worksheet.write_string(5, 0, "#DAA520")?;
    worksheet.write_blank(5, 1, &color_format)?;

    // Add a Theme color.
    let color_format = Format::new().set_background_color(Color::Theme(4, 3));
    worksheet.write_string(6, 0, "Theme(4, 3)")?;
    worksheet.write_blank(6, 1, &color_format)?;

    // Save the file to disk.
    workbook.save("into_color.xlsx")?;

    Ok(())
}
