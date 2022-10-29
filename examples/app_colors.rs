// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! A sample palette of the the defined colors and user defined RGB colors
//! available in the rust_xlsxwriter library.

use rust_xlsxwriter::*;

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    // Write some enum defined colors to cells.
    let color_format = Format::new().set_background_color(XlsxColor::Black);
    worksheet.write_string_only(0, 0, "Black")?;
    worksheet.write_blank(0, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::Blue);
    worksheet.write_string_only(1, 0, "Blue")?;
    worksheet.write_blank(1, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::Brown);
    worksheet.write_string_only(2, 0, "Brown")?;
    worksheet.write_blank(2, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::Cyan);
    worksheet.write_string_only(3, 0, "Cyan")?;
    worksheet.write_blank(3, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::Gray);
    worksheet.write_string_only(4, 0, "Gray")?;
    worksheet.write_blank(4, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::Green);
    worksheet.write_string_only(5, 0, "Green")?;
    worksheet.write_blank(5, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::Lime);
    worksheet.write_string_only(6, 0, "Lime")?;
    worksheet.write_blank(6, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::Magenta);
    worksheet.write_string_only(7, 0, "Magenta")?;
    worksheet.write_blank(7, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::Navy);
    worksheet.write_string_only(8, 0, "Navy")?;
    worksheet.write_blank(8, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::Orange);
    worksheet.write_string_only(9, 0, "Orange")?;
    worksheet.write_blank(9, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::Pink);
    worksheet.write_string_only(10, 0, "Pink")?;
    worksheet.write_blank(10, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::Purple);
    worksheet.write_string_only(11, 0, "Purple")?;
    worksheet.write_blank(11, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::Red);
    worksheet.write_string_only(12, 0, "Red")?;
    worksheet.write_blank(12, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::Silver);
    worksheet.write_string_only(13, 0, "Silver")?;
    worksheet.write_blank(13, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::White);
    worksheet.write_string_only(14, 0, "White")?;
    worksheet.write_blank(14, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::Yellow);
    worksheet.write_string_only(15, 0, "Yellow")?;
    worksheet.write_blank(15, 1, &color_format)?;

    // Write some user defined RGB colors to cells.
    let color_format = Format::new().set_background_color(XlsxColor::RGB(0xFF7F50));
    worksheet.write_string_only(16, 0, "#FF7F50")?;
    worksheet.write_blank(16, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::RGB(0xDCDCDC));
    worksheet.write_string_only(17, 0, "#DCDCDC")?;
    worksheet.write_blank(17, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::RGB(0x6495ED));
    worksheet.write_string_only(18, 0, "#6495ED")?;
    worksheet.write_blank(18, 1, &color_format)?;

    let color_format = Format::new().set_background_color(XlsxColor::RGB(0xDAA520));
    worksheet.write_string_only(19, 0, "#DAA520")?;
    worksheet.write_blank(19, 1, &color_format)?;

    // Save the file to disk.
    workbook.save("colors.xlsx")?;

    Ok(())
}
