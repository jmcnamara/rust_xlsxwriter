// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! A simple example of setting some "freeze" panes in worksheets using the
//! rust_xlsxwriter library.

use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create some formats to use in the worksheet.
    let header_format = Format::new()
        .set_bold()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_foreground_color(Color::RGB(0xD7E4BC))
        .set_border(FormatBorder::Thin);

    let center_format = Format::new().set_align(FormatAlign::Center);

    // Some range limits to use in this example.
    let max_row = 50;
    let max_col = 26;

    // -----------------------------------------------------------------------
    // Example 1. Freeze pane on the top row.
    // -----------------------------------------------------------------------
    let worksheet1 = workbook.add_worksheet().set_name("Panes 1")?;

    // Freeze the top row only.
    worksheet1.set_freeze_panes(1, 0)?;

    // Add some data and formatting to the worksheet.
    worksheet1.set_row_height(0, 20)?;
    for col in 0..max_col {
        worksheet1.write_string_with_format(0, col, "Scroll down", &header_format)?;
        worksheet1.set_column_width(col, 16)?;
    }
    for row in 1..max_row {
        for col in 0..max_col {
            worksheet1.write_number_with_format(row, col, row + 1, &center_format)?;
        }
    }

    // -----------------------------------------------------------------------
    // Example 2. Freeze pane on the left column.
    // -----------------------------------------------------------------------
    let worksheet2 = workbook.add_worksheet().set_name("Panes 2")?;

    // Freeze the leftmost column only.
    worksheet2.set_freeze_panes(0, 1)?;

    // Add some data and formatting to the worksheet.
    worksheet2.set_column_width(0, 16)?;
    for row in 0..max_row {
        worksheet2.write_string_with_format(row, 0, "Scroll Across", &header_format)?;

        for col in 1..max_col {
            worksheet2.write_number_with_format(row, col, col, &center_format)?;
        }
    }

    // -----------------------------------------------------------------------
    // Example 3. Freeze pane on the top row and leftmost column.
    // -----------------------------------------------------------------------
    let worksheet3 = workbook.add_worksheet().set_name("Panes 3")?;

    // Freeze the top row and leftmost column.
    worksheet3.set_freeze_panes(1, 1)?;

    // Add some data and formatting to the worksheet.
    worksheet3.set_row_height(0, 20)?;
    worksheet3.set_column_width(0, 16)?;
    worksheet3.write_blank(0, 0, &header_format)?;

    for col in 1..max_col {
        worksheet3.write_string_with_format(0, col, "Scroll down", &header_format)?;
        worksheet3.set_column_width(col, 16)?;
    }
    for row in 1..max_row {
        worksheet3.write_string_with_format(row, 0, "Scroll Across", &header_format)?;

        for col in 1..max_col {
            worksheet3.write_number_with_format(row, col, col, &center_format)?;
        }
    }

    // -----------------------------------------------------------------------
    // Example 4. Freeze pane on the top row and leftmost column, with
    //            scrolling area shifted.
    // -----------------------------------------------------------------------
    let worksheet4 = workbook.add_worksheet().set_name("Panes 4")?;

    // Freeze the top row and leftmost column.
    worksheet4.set_freeze_panes(1, 1)?;

    // Shift the scrolled area in the scrolling pane.
    worksheet4.set_freeze_panes_top_cell(20, 12)?;

    // Add some data and formatting to the worksheet.
    worksheet4.set_row_height(0, 20)?;
    worksheet4.set_column_width(0, 16)?;
    worksheet4.write_blank(0, 0, &header_format)?;

    for col in 1..max_col {
        worksheet4.write_string_with_format(0, col, "Scroll down", &header_format)?;
        worksheet4.set_column_width(col, 16)?;
    }
    for row in 1..max_row {
        worksheet4.write_string_with_format(row, 0, "Scroll Across", &header_format)?;

        for col in 1..max_col {
            worksheet4.write_number_with_format(row, col, col, &center_format)?;
        }
    }

    // Save the file to disk.
    workbook.save("panes.xlsx")?;

    Ok(())
}
