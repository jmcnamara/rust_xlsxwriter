// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of how to group columns into outlines with the `rust_xlsxwriter`
//! library.
//!
//! In Excel an outline is a group of rows or columns that can be collapsed or
//! expanded to simplify hierarchical data. It is often used with the
//! `SUBTOTAL()` function.

use rust_xlsxwriter::{Format, Workbook, Worksheet, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a format to use in headings.
    let bold = Format::new().set_bold();

    // -----------------------------------------------------------------------
    // 1. Add an outline column group with sub-total.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Simple outline column grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add grouping for the over the sub-total range.
    worksheet.group_columns(1, 8)?;

    // -----------------------------------------------------------------------
    // 2. Add nested outline column groups with sub-totals.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Nested outline column grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add grouping for the over the sub-total range.
    worksheet.group_columns(1, 8)?;

    // Add secondary groups within the first range.
    worksheet.group_columns(1, 3)?;
    worksheet.group_columns(5, 7)?;

    // -----------------------------------------------------------------------
    // 3. Add a collapsed inner outline column groups.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Collapsed inner column grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add grouping for the over the sub-total range.
    worksheet.group_columns(1, 8)?;

    // Add collapsed secondary groups within the first range.
    worksheet.group_columns_collapsed(1, 3)?;
    worksheet.group_columns_collapsed(5, 7)?;

    // -----------------------------------------------------------------------
    // 4. Add a collapsed outer column group.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Collapsed outer column grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add collapsed grouping for the over the sub-total range.
    worksheet.group_columns_collapsed(1, 8)?;

    // Add secondary groups within the first range.
    worksheet.group_columns(1, 3)?;
    worksheet.group_columns(5, 7)?;

    // -----------------------------------------------------------------------
    // 5. Column groups with outline symbols on top.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Outline column grouping symbols to the left.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add outline column groups.
    worksheet.group_columns(1, 3)?;
    worksheet.group_columns(5, 7)?;

    // Change the worksheet group setting so outline symbols are to the left.
    worksheet.group_symbols_to_left(true);

    // -----------------------------------------------------------------------
    // 6. Demonstrate all group levels.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Excel outline levels.";
    let levels = [
        "Level 1", "Level 2", "Level 3", "Level 4", //
        "Level 5", "Level 6", "Level 7", "Level 6", //
        "Level 5", "Level 4", "Level 3", "Level 2", //
        "Level 1",
    ];
    worksheet.write_row(0, 0, levels)?;
    worksheet.write_with_format(2, 0, description, &bold)?;

    // Add outline column groups from outer to inner.
    worksheet.group_columns(0, 12)?;
    worksheet.group_columns(1, 11)?;
    worksheet.group_columns(2, 10)?;
    worksheet.group_columns(3, 9)?;
    worksheet.group_columns(4, 8)?;
    worksheet.group_columns(5, 7)?;
    worksheet.group_columns(6, 6)?;

    // Save the file to disk.
    workbook.save("grouped_columns.xlsx")?;

    Ok(())
}

// Generate worksheet data.
pub fn populate_worksheet_data(
    worksheet: &mut Worksheet,
    description: &str,
    bold: &Format,
) -> Result<(), XlsxError> {
    worksheet.write_with_format(0, 0, "Region", bold)?;
    worksheet.write_with_format(1, 0, "North", bold)?;
    worksheet.write_with_format(2, 0, "South", bold)?;
    worksheet.write_with_format(3, 0, "East", bold)?;
    worksheet.write_with_format(4, 0, "West", bold)?;

    worksheet.write_with_format(0, 1, "Jan", bold)?;
    worksheet.write(1, 1, 50)?;
    worksheet.write(2, 1, 10)?;
    worksheet.write(3, 1, 45)?;
    worksheet.write(4, 1, 15)?;

    worksheet.write_with_format(0, 2, "Feb", bold)?;
    worksheet.write(1, 2, 20)?;
    worksheet.write(2, 2, 20)?;
    worksheet.write(3, 2, 75)?;
    worksheet.write(4, 2, 15)?;

    worksheet.write_with_format(0, 3, "Mar", bold)?;
    worksheet.write(1, 3, 15)?;
    worksheet.write(2, 3, 30)?;
    worksheet.write(3, 3, 50)?;
    worksheet.write(4, 3, 35)?;

    worksheet.write_with_format(0, 4, "Q1 Total", bold)?;
    worksheet.write_formula_with_format(1, 4, "=SUBTOTAL(9,B2:D2)", bold)?;
    worksheet.write_formula_with_format(2, 4, "=SUBTOTAL(9,B3:D3)", bold)?;
    worksheet.write_formula_with_format(3, 4, "=SUBTOTAL(9,B4:D4)", bold)?;
    worksheet.write_formula_with_format(4, 4, "=SUBTOTAL(9,B5:D5)", bold)?;

    worksheet.write_with_format(0, 5, "Apr", bold)?;
    worksheet.write(1, 5, 25)?;
    worksheet.write(2, 5, 50)?;
    worksheet.write(3, 5, 15)?;
    worksheet.write(4, 5, 35)?;

    worksheet.write_with_format(0, 6, "May", bold)?;
    worksheet.write(1, 6, 65)?;
    worksheet.write(2, 6, 50)?;
    worksheet.write(3, 6, 75)?;
    worksheet.write(4, 6, 70)?;

    worksheet.write_with_format(0, 7, "Jun", bold)?;
    worksheet.write(1, 7, 80)?;
    worksheet.write(2, 7, 50)?;
    worksheet.write(3, 7, 90)?;
    worksheet.write(4, 7, 50)?;

    worksheet.write_with_format(0, 8, "Q2 Total", bold)?;
    worksheet.write_formula_with_format(1, 8, "=SUBTOTAL(9,F2:H2)", bold)?;
    worksheet.write_formula_with_format(2, 8, "=SUBTOTAL(9,F3:H3)", bold)?;
    worksheet.write_formula_with_format(3, 8, "=SUBTOTAL(9,F4:H4)", bold)?;
    worksheet.write_formula_with_format(4, 8, "=SUBTOTAL(9,F5:H5)", bold)?;

    worksheet.write_with_format(0, 9, "H1 Total", bold)?;
    worksheet.write_formula_with_format(1, 9, "=SUBTOTAL(9,B2:I2)", bold)?;
    worksheet.write_formula_with_format(2, 9, "=SUBTOTAL(9,B3:I3)", bold)?;
    worksheet.write_formula_with_format(3, 9, "=SUBTOTAL(9,B4:I4)", bold)?;
    worksheet.write_formula_with_format(4, 9, "=SUBTOTAL(9,B5:I5)", bold)?;

    // Autofit the columns for clarity.
    worksheet.autofit();

    worksheet.write_with_format(6, 0, description, bold)?;

    Ok(())
}
