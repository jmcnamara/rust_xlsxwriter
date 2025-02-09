// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! This example demonstrates how cells without explicit formats inherit the
//! formats from the row and column that they are in. Note the output:
//!
//! - Cell C1 has a green font color.
//! - Cell A3 has a bold format.
//! - Cell C3 has both a bold format and a green font color.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    // Add some formats.
    let red = Format::new().set_font_color("9C0006");
    let bold = Format::new().set_bold();

    // Set some row and column formats.
    worksheet.set_row_format(2, &bold)?;
    worksheet.set_column_format(2, &red)?;

    // Write some strings without explicit formats.
    worksheet.write(0, 2, "C1")?; // Red.
    worksheet.write(2, 0, "A3")?; // Bold.
    worksheet.write(2, 2, "C3")?; // Bold and red.

    // Save the file.
    workbook.save("formats.xlsx")?;

    Ok(())
}
