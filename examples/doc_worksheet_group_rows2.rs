// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of how to group worksheet rows into outlines. This example shows
//! hows to add secondary groups within a primary grouping. Excel requires at
//! least one row between each outline grouping at the same level.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a format to use in headings.
    let bold = Format::new().set_bold();

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();

    worksheet.write_with_format(0, 0, "Region", &bold)?;
    worksheet.write(1, 0, "North 1")?;
    worksheet.write(2, 0, "North 2")?;
    worksheet.write(3, 0, "North 3")?;
    worksheet.write(4, 0, "North 4")?;
    worksheet.write_with_format(5, 0, "North Total", &bold)?;

    worksheet.write_with_format(0, 1, "Sales", &bold)?;
    worksheet.write(1, 1, 1000)?;
    worksheet.write(2, 1, 1200)?;
    worksheet.write(3, 1, 900)?;
    worksheet.write(4, 1, 1200)?;
    worksheet.write_formula_with_format(5, 1, "=SUBTOTAL(9,B2:B5)", &bold)?;

    worksheet.write(6, 0, "South 1")?;
    worksheet.write(7, 0, "South 2")?;
    worksheet.write(8, 0, "South 3")?;
    worksheet.write(9, 0, "South 4")?;
    worksheet.write_with_format(10, 0, "South Total", &bold)?;

    worksheet.write(6, 1, 400)?;
    worksheet.write(7, 1, 600)?;
    worksheet.write(8, 1, 500)?;
    worksheet.write(9, 1, 600)?;
    worksheet.write_formula_with_format(10, 1, "=SUBTOTAL(9,B7:B10)", &bold)?;

    worksheet.write_with_format(11, 0, "Grand Total", &bold)?;
    worksheet.write_formula_with_format(11, 1, "=SUBTOTAL(9,B2:B11)", &bold)?;

    // Autofit the columns for clarity.
    worksheet.autofit();

    // Add grouping over the sub-total range.
    worksheet.group_rows(1, 10)?;

    // Add secondary groups within the first range.
    worksheet.group_rows(1, 4)?;
    worksheet.group_rows(6, 9)?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
