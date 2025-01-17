// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of the various cell formatting options that are available in the
//! rust_xlsxwriter library. These are laid out on worksheets that correspond to
//! the sections of the Excel "Format Cells" dialog.

use rust_xlsxwriter::*;

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a general heading format.
    let header_format = Format::new()
        .set_bold()
        .set_align(FormatAlign::Top)
        .set_border(FormatBorder::Thin)
        .set_background_color(Color::RGB(0xC6EFCE));

    // -----------------------------------------------------------------------
    // Create a worksheet that demonstrates number formatting.
    // -----------------------------------------------------------------------
    let worksheet = workbook.add_worksheet().set_name("Number")?;

    // Make the header row and columns taller and wider for clarity.
    worksheet.set_column_width(0, 18)?;
    worksheet.set_column_width(1, 18)?;
    worksheet.set_row_height(0, 25)?;

    // Add some descriptive text and the formatted numbers.
    worksheet.write_string_with_format(0, 0, "Number Categories", &header_format)?;
    worksheet.write_string_with_format(0, 1, "Formatted Numbers", &header_format)?;

    // Write an unformatted number with the default or "General" format.
    worksheet.write_string(1, 0, "General")?;
    worksheet.write_number(1, 1, 1234.567)?;

    // Write a number with a decimal format.
    worksheet.write_string(2, 0, "Number")?;
    let decimal_format = Format::new().set_num_format("0.00");
    worksheet.write_number_with_format(2, 1, 1234.567, &decimal_format)?;

    // Write a number with a currency format.
    worksheet.write_string(3, 0, "Currency")?;
    let currency_format = Format::new().set_num_format("[$¥-ja-JP]#,##0.00");
    worksheet.write_number_with_format(3, 1, 1234.567, &currency_format)?;

    // Write a number with an accountancy format.
    worksheet.write_string(4, 0, "Accountancy")?;
    let accountancy_format = Format::new().set_num_format("_-[$¥-ja-JP]* #,##0.00_-");
    worksheet.write_number_with_format(4, 1, 1234.567, &accountancy_format)?;

    // Write a number with a short date format.
    worksheet.write_string(5, 0, "Date")?;
    let short_date_format = Format::new().set_num_format("yyyy-mm-dd;@");
    worksheet.write_number_with_format(5, 1, 44927.23, &short_date_format)?;

    // Write a number with a long date format.
    worksheet.write_string(6, 0, "Date")?;
    let long_date_format = Format::new().set_num_format("[$-x-sysdate]dddd, mmmm dd, yyyy");
    worksheet.write_number_with_format(6, 1, 44927.23, &long_date_format)?;

    // Write a number with a percentage format.
    worksheet.write_string(7, 0, "Percentage")?;
    let percentage_format = Format::new().set_num_format("0.00%");
    worksheet.write_number_with_format(7, 1, 72.5 / 100.0, &percentage_format)?;

    // Write a number with a fraction format.
    worksheet.write_string(8, 0, "Fraction")?;
    let fraction_format = Format::new().set_num_format("# ??/??");
    worksheet.write_number_with_format(8, 1, 5.0 / 16.0, &fraction_format)?;

    // Write a number with a percentage format.
    worksheet.write_string(9, 0, "Scientific")?;
    let scientific_format = Format::new().set_num_format("0.00E+00");
    worksheet.write_number_with_format(9, 1, 1234.567, &scientific_format)?;

    // Write a number with a text format.
    worksheet.write_string(10, 0, "Text")?;
    let text_format = Format::new().set_num_format("@");
    worksheet.write_number_with_format(10, 1, 1234.567, &text_format)?;

    // -----------------------------------------------------------------------
    // Create a worksheet that demonstrates number formatting.
    // -----------------------------------------------------------------------
    let worksheet = workbook.add_worksheet().set_name("Alignment")?;

    // Make some rows and columns taller and wider for clarity.
    worksheet.set_column_width(0, 18)?;
    for row_num in 0..5 {
        worksheet.set_row_height(row_num, 30)?;
    }

    // Add some descriptive text at the top of the worksheet.
    worksheet.write_string_with_format(0, 0, "Alignment formats", &header_format)?;

    // Some examples of positional alignment formats.
    let center_format = Format::new().set_align(FormatAlign::Center);
    worksheet.write_string_with_format(1, 0, "Center", &center_format)?;

    let top_left_format = Format::new()
        .set_align(FormatAlign::Top)
        .set_align(FormatAlign::Left);
    worksheet.write_string_with_format(2, 0, "Top - Left", &top_left_format)?;

    let center_center_format = Format::new()
        .set_align(FormatAlign::VerticalCenter)
        .set_align(FormatAlign::Center);
    worksheet.write_string_with_format(3, 0, "Center - Center", &center_center_format)?;

    let bottom_right_format = Format::new()
        .set_align(FormatAlign::Bottom)
        .set_align(FormatAlign::Right);
    worksheet.write_string_with_format(4, 0, "Bottom - Right", &bottom_right_format)?;

    // Some indentation formats.
    let indent1_format = Format::new().set_indent(1);
    worksheet.write_string_with_format(5, 0, "Indent 1", &indent1_format)?;

    let indent2_format = Format::new().set_indent(2);
    worksheet.write_string_with_format(6, 0, "Indent 2", &indent2_format)?;

    // Text wrap format.
    let text_wrap_format = Format::new().set_text_wrap();
    worksheet.write_string_with_format(7, 0, "Some text that is wrapped", &text_wrap_format)?;
    worksheet.write_string_with_format(8, 0, "Text\nwrapped\nat newlines", &text_wrap_format)?;

    // Shrink text format.
    let shrink_format = Format::new().set_shrink();
    worksheet.write_string_with_format(9, 0, "Shrink wide text to fit cell", &shrink_format)?;

    // Text rotation formats.
    let rotate_format1 = Format::new().set_rotation(30);
    worksheet.write_string_with_format(10, 0, "Rotate", &rotate_format1)?;

    let rotate_format2 = Format::new().set_rotation(-30);
    worksheet.write_string_with_format(11, 0, "Rotate", &rotate_format2)?;

    let rotate_format3 = Format::new().set_rotation(270);
    worksheet.write_string_with_format(12, 0, "Rotate", &rotate_format3)?;

    // -----------------------------------------------------------------------
    // Create a worksheet that demonstrates font formatting.
    // -----------------------------------------------------------------------
    let worksheet = workbook.add_worksheet().set_name("Font")?;

    // Make the header row and columns taller and wider for clarity.
    worksheet.set_column_width(0, 18)?;
    worksheet.set_column_width(1, 18)?;
    worksheet.set_row_height(0, 25)?;

    // Add some descriptive text at the top of the worksheet.
    worksheet.write_string_with_format(0, 0, "Font formatting", &header_format)?;

    // Different fonts.
    worksheet.write_string(1, 0, "Calibri 11 (default font)")?;

    let algerian_format = Format::new().set_font_name("Algerian");
    worksheet.write_string_with_format(2, 0, "Algerian", &algerian_format)?;

    let consolas_format = Format::new().set_font_name("Consolas");
    worksheet.write_string_with_format(3, 0, "Consolas", &consolas_format)?;

    let comic_sans_format = Format::new().set_font_name("Comic Sans MS");
    worksheet.write_string_with_format(4, 0, "Comic Sans MS", &comic_sans_format)?;

    // Font styles.
    let bold = Format::new().set_bold();
    worksheet.write_string_with_format(5, 0, "Bold", &bold)?;

    let italic = Format::new().set_italic();
    worksheet.write_string_with_format(6, 0, "Italic", &italic)?;

    let bold_italic = Format::new().set_bold().set_italic();
    worksheet.write_string_with_format(7, 0, "Bold/Italic", &bold_italic)?;

    // Font size.
    let size_format = Format::new().set_font_size(18);
    worksheet.write_string_with_format(8, 0, "Font size 18", &size_format)?;

    // Font color.
    let font_color_format = Format::new().set_font_color(Color::Red);
    worksheet.write_string_with_format(9, 0, "Font color", &font_color_format)?;

    // Font underline.
    let underline_format = Format::new().set_underline(FormatUnderline::Single);
    worksheet.write_string_with_format(10, 0, "Underline", &underline_format)?;

    // Font strike-though.
    let strikethrough_format = Format::new().set_font_strikethrough();
    worksheet.write_string_with_format(11, 0, "Strikethrough", &strikethrough_format)?;

    // -----------------------------------------------------------------------
    // Create a worksheet that demonstrates border formatting.
    // -----------------------------------------------------------------------
    let worksheet = workbook.add_worksheet().set_name("Border")?;

    // Make the header row and columns taller and wider for clarity.
    worksheet.set_column_width(2, 18)?;
    worksheet.set_row_height(0, 25)?;

    // Add some descriptive text at the top of the worksheet.
    worksheet.write_string_with_format(0, 2, "Border formats", &header_format)?;

    // Add some borders to cells.
    let border_format1 = Format::new().set_border(FormatBorder::Thin);
    worksheet.write_string_with_format(2, 2, "Thin Border", &border_format1)?;

    let border_format2 = Format::new().set_border(FormatBorder::Dotted);
    worksheet.write_string_with_format(4, 2, "Dotted Border", &border_format2)?;

    let border_format3 = Format::new().set_border(FormatBorder::Double);
    worksheet.write_string_with_format(6, 2, "Double Border", &border_format3)?;

    let border_format4 = Format::new()
        .set_border(FormatBorder::Thin)
        .set_border_color(Color::Red);
    worksheet.write_string_with_format(8, 2, "Color Border", &border_format4)?;

    // -----------------------------------------------------------------------
    // Create a worksheet that demonstrates fill/pattern formatting.
    // -----------------------------------------------------------------------
    let worksheet = workbook.add_worksheet().set_name("Fill")?;

    // Make the header row and columns taller and wider for clarity.
    worksheet.set_column_width(1, 18)?;
    worksheet.set_row_height(0, 25)?;

    // Add some descriptive text at the top of the worksheet.
    worksheet.write_string_with_format(0, 1, "Fill formats", &header_format)?;

    // Write some cells with pattern fills.
    let fill_format1 = Format::new()
        .set_background_color(Color::Yellow)
        .set_pattern(FormatPattern::Solid);
    worksheet.write_string_with_format(2, 1, "Solid fill", &fill_format1)?;

    let fill_format2 = Format::new()
        .set_background_color(Color::Yellow)
        .set_foreground_color(Color::Orange)
        .set_pattern(FormatPattern::Gray0625);
    worksheet.write_string_with_format(4, 1, "Pattern fill", &fill_format2)?;

    // Save the file to disk.
    workbook.save("cell_formats.xlsx")?;

    Ok(())
}
