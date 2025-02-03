// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of how to group worksheet columns into outlines. This example
//! puts the expand/collapse symbol to the left of the range for all row groups
//! in the worksheet.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a format to use in headings.
    let bold = Format::new().set_bold();

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();

    worksheet.write_with_format(0, 0, "Region", &bold)?;
    worksheet.write_with_format(1, 0, "North", &bold)?;
    worksheet.write_with_format(2, 0, "South", &bold)?;
    worksheet.write_with_format(3, 0, "East", &bold)?;
    worksheet.write_with_format(4, 0, "West", &bold)?;

    worksheet.write_with_format(0, 1, "Jan", &bold)?;
    worksheet.write(1, 1, 50)?;
    worksheet.write(2, 1, 10)?;
    worksheet.write(3, 1, 45)?;
    worksheet.write(4, 1, 15)?;

    worksheet.write_with_format(0, 2, "Feb", &bold)?;
    worksheet.write(1, 2, 20)?;
    worksheet.write(2, 2, 20)?;
    worksheet.write(3, 2, 75)?;
    worksheet.write(4, 2, 15)?;

    worksheet.write_with_format(0, 3, "Mar", &bold)?;
    worksheet.write(1, 3, 15)?;
    worksheet.write(2, 3, 30)?;
    worksheet.write(3, 3, 50)?;
    worksheet.write(4, 3, 35)?;

    worksheet.write_with_format(0, 4, "Q1 Total", &bold)?;
    worksheet.write_formula_with_format(1, 4, "=SUBTOTAL(9,B2:D2)", &bold)?;
    worksheet.write_formula_with_format(2, 4, "=SUBTOTAL(9,B3:D3)", &bold)?;
    worksheet.write_formula_with_format(3, 4, "=SUBTOTAL(9,B4:D4)", &bold)?;
    worksheet.write_formula_with_format(4, 4, "=SUBTOTAL(9,B5:D5)", &bold)?;

    worksheet.write_with_format(0, 5, "Apr", &bold)?;
    worksheet.write(1, 5, 25)?;
    worksheet.write(2, 5, 50)?;
    worksheet.write(3, 5, 15)?;
    worksheet.write(4, 5, 35)?;

    worksheet.write_with_format(0, 6, "May", &bold)?;
    worksheet.write(1, 6, 65)?;
    worksheet.write(2, 6, 50)?;
    worksheet.write(3, 6, 75)?;
    worksheet.write(4, 6, 70)?;

    worksheet.write_with_format(0, 7, "Jun", &bold)?;
    worksheet.write(1, 7, 80)?;
    worksheet.write(2, 7, 50)?;
    worksheet.write(3, 7, 90)?;
    worksheet.write(4, 7, 50)?;

    worksheet.write_with_format(0, 8, "Q2 Total", &bold)?;
    worksheet.write_formula_with_format(1, 8, "=SUBTOTAL(9,F2:H2)", &bold)?;
    worksheet.write_formula_with_format(2, 8, "=SUBTOTAL(9,F3:H3)", &bold)?;
    worksheet.write_formula_with_format(3, 8, "=SUBTOTAL(9,F4:H4)", &bold)?;
    worksheet.write_formula_with_format(4, 8, "=SUBTOTAL(9,F5:H5)", &bold)?;

    // Autofit the columns for clarity.
    worksheet.autofit();

    // Add groupings for the over the sub-total ranges.
    worksheet.group_columns(1, 3)?;
    worksheet.group_columns(5, 7)?;

    // Display the expand/collapse symbol to the left of the range.
    worksheet.group_symbols_to_left(true);

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
