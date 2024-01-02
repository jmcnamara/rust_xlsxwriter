// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of using the rust_xlsxwriter library to write "rich" multi-format
//! strings in worksheet cells.

use rust_xlsxwriter::{Color, Format, FormatAlign, FormatScript, Workbook, XlsxError};

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
    let bold = Format::new().set_bold();
    let italic = Format::new().set_italic();
    let center = Format::new().set_align(FormatAlign::Center);
    let superscript = Format::new().set_font_script(FormatScript::Superscript);

    // Write some rich strings with multiple formats.
    let segments = [
        (&default, "This is "),
        (&bold, "bold"),
        (&default, " and this is "),
        (&italic, "italic"),
    ];
    worksheet.write_rich_string(0, 0, &segments)?;

    let segments = [
        (&default, "This is "),
        (&red, "red"),
        (&default, " and this is "),
        (&blue, "blue"),
    ];
    worksheet.write_rich_string(2, 0, &segments)?;

    let segments = [
        (&default, "Some "),
        (&bold, "bold text"),
        (&default, " centered"),
    ];
    worksheet.write_rich_string_with_format(4, 0, &segments, &center)?;

    let segments = [(&italic, "j = k"), (&superscript, "(n-1)")];
    worksheet.write_rich_string_with_format(6, 0, &segments, &center)?;

    // It is possible, and idiomatic, to use slices as the string segments.
    let text = "This is blue and this is red";
    let segments = [
        (&default, &text[..8]),
        (&blue, &text[8..12]),
        (&default, &text[12..25]),
        (&red, &text[25..]),
    ];
    worksheet.write_rich_string(8, 0, &segments)?;

    // Save the file to disk.
    workbook.save("rich_strings.xlsx")?;

    Ok(())
}
