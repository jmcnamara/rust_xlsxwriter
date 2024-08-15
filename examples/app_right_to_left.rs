// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of using rust_xlsxwriter to create a workbook with the default
//! worksheet and cell text direction changed from left-to-right to
//! right-to-left, as required by some middle eastern versions of Excel.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add the cell formats.
    let format_left_to_right = Format::new().set_reading_direction(1);
    let format_right_to_left = Format::new().set_reading_direction(2);

    // Add a worksheet in the standard left to right direction.
    let worksheet1 = workbook.add_worksheet();

    // Make the column wider for clarity.
    worksheet1.set_column_width(0, 25)?;

    // Standard direction:         | A1 | B1 | C1 | ...
    worksheet1.write(0, 0, "نص عربي / English text")?;
    worksheet1.write_with_format(1, 0, "نص عربي / English text", &format_left_to_right)?;
    worksheet1.write_with_format(2, 0, "نص عربي / English text", &format_right_to_left)?;

    // Add a worksheet and change it to right to left direction.
    let worksheet2 = workbook.add_worksheet();
    worksheet2.set_right_to_left(true);

    // Make the column wider for clarity.
    worksheet2.set_column_width(0, 25)?;

    // Right to left direction:    ... | C1 | B1 | A1 |
    worksheet2.write(0, 0, "نص عربي / English text")?;
    worksheet2.write_with_format(1, 0, "نص عربي / English text", &format_left_to_right)?;
    worksheet2.write_with_format(2, 0, "نص عربي / English text", &format_right_to_left)?;

    workbook.save("right_to_left.xlsx")?;

    Ok(())
}
