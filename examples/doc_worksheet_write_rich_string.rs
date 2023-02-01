// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing a "rich" string with multiple
//! formats, and an additional cell format.

use rust_xlsxwriter::{Format, FormatAlign, Workbook, XlsxColor, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    worksheet.set_column_width(0, 30)?;

    // Add some formats to use in the rich strings.
    let default = Format::default();
    let red = Format::new().set_font_color(XlsxColor::Red);
    let blue = Format::new().set_font_color(XlsxColor::Blue);

    // Write a rich strings with multiple formats.
    let segments = [
        (&default, "This is "),
        (&red, "red"),
        (&default, " and this is "),
        (&blue, "blue"),
    ];
    worksheet.write_rich_string_only(0, 0, &segments)?;

    // Add an extra format to use for the entire cell.
    let center = Format::new().set_align(FormatAlign::Center);

    // Write the rich string again with the cell format.
    worksheet.write_rich_string(2, 0, &segments, &center)?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
