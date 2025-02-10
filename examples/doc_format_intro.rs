// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates some of the available formatting
//! properties.

use rust_xlsxwriter::{Color, Format, FormatBorder, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    // Make the first column wider for clarity.
    worksheet.set_column_width(0, 14)?;

    // Create some sample formats to display
    let format1 = Format::new().set_font_name("Arial");
    worksheet.write_with_format(0, 0, "Fonts", &format1)?;

    let format2 = Format::new().set_font_name("Algerian").set_font_size(14);
    worksheet.write_with_format(1, 0, "Fonts", &format2)?;

    let format3 = Format::new().set_font_name("Comic Sans MS");
    worksheet.write_with_format(2, 0, "Fonts", &format3)?;

    let format4 = Format::new().set_font_name("Edwardian Script ITC");
    worksheet.write_with_format(3, 0, "Fonts", &format4)?;

    let format5 = Format::new().set_font_color(Color::Red);
    worksheet.write_with_format(4, 0, "Font color", &format5)?;

    let format6 = Format::new().set_background_color(Color::RGB(0xDAA520));
    worksheet.write_with_format(5, 0, "Fills", &format6)?;

    let format7 = Format::new().set_border(FormatBorder::Thin);
    worksheet.write_with_format(6, 0, "Borders", &format7)?;

    let format8 = Format::new().set_bold();
    worksheet.write_with_format(7, 0, "Bold", &format8)?;

    let format9 = Format::new().set_italic();
    worksheet.write_with_format(8, 0, "Italic", &format9)?;

    let format10 = Format::new().set_bold().set_italic();
    worksheet.write_with_format(9, 0, "Bold and Italic", &format10)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
