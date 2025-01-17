// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing a "rich" string with multiple
//! formats.

use rust_xlsxwriter::{Color, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    worksheet.set_column_width(0, 30)?;

    // Add some formats to use in the rich strings.
    let default = Format::default();
    let red = Format::new().set_font_color(Color::Red);
    let blue = Format::new().set_font_color(Color::Blue);

    // Write a rich strings with multiple formats.
    let segments = [
        (&default, "This is "),
        (&red, "red"),
        (&default, " and this is "),
        (&blue, "blue"),
    ];
    worksheet.write_rich_string(0, 0, &segments)?;

    // It is possible, and idiomatic, to use slices as the string segments.
    let text = "This is blue and this is red";
    let segments = [
        (&default, &text[..8]),
        (&blue, &text[8..12]),
        (&default, &text[12..25]),
        (&red, &text[25..]),
    ];
    worksheet.write_rich_string(1, 0, &segments)?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
