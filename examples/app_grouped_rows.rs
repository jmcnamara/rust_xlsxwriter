// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of how to group rows into outlines with the `rust_xlsxwriter`
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
    // 1. Add an outline row group with sub-total.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Simple outline row grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add grouping for the over the sub-total range.
    worksheet.group_rows(1, 10)?;

    // -----------------------------------------------------------------------
    // 2. Add nested outline row groups with sub-totals.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Nested outline row grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add grouping for the over the sub-total range.
    worksheet.group_rows(1, 10)?;

    // Add secondary groups within the first range.
    worksheet.group_rows(1, 4)?;
    worksheet.group_rows(6, 9)?;

    // -----------------------------------------------------------------------
    // 3. Add a collapsed inner outline row groups.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Collapsed inner row grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add grouping for the over the sub-total range.
    worksheet.group_rows(1, 10)?;

    // Add collapsed secondary groups within the first range.
    worksheet.group_rows_collapsed(1, 4)?;
    worksheet.group_rows_collapsed(6, 9)?;

    // -----------------------------------------------------------------------
    // 4. Add a collapsed outer row group.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Collapsed outer row grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add collapsed grouping for the over the sub-total range.
    worksheet.group_rows_collapsed(1, 10)?;

    // Add secondary groups within the first range.
    worksheet.group_rows(1, 4)?;
    worksheet.group_rows(6, 9)?;

    // -----------------------------------------------------------------------
    // 5. Row groups with outline symbols on top.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Outline row grouping symbols on top.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add outline row groups.
    worksheet.group_rows(1, 4)?;
    worksheet.group_rows(6, 9)?;

    // Change the worksheet group setting so the outline symbols are on top.
    worksheet.group_symbols_above(true);

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
    worksheet.write_column(0, 0, levels)?;
    worksheet.write_with_format(0, 3, description, &bold)?;

    // Add outline row groups from outer to inner.
    worksheet.group_rows(0, 12)?;
    worksheet.group_rows(1, 11)?;
    worksheet.group_rows(2, 10)?;
    worksheet.group_rows(3, 9)?;
    worksheet.group_rows(4, 8)?;
    worksheet.group_rows(5, 7)?;
    worksheet.group_rows(6, 6)?;

    // Save the file to disk.
    workbook.save("grouped_rows.xlsx")?;

    Ok(())
}

// Generate worksheet data.
pub fn populate_worksheet_data(
    worksheet: &mut Worksheet,
    description: &str,
    bold: &Format,
) -> Result<(), XlsxError> {
    worksheet.write_with_format(0, 3, description, bold)?;

    worksheet.write_with_format(0, 0, "Region", bold)?;
    worksheet.write(1, 0, "North 1")?;
    worksheet.write(2, 0, "North 2")?;
    worksheet.write(3, 0, "North 3")?;
    worksheet.write(4, 0, "North 4")?;
    worksheet.write_with_format(5, 0, "North Total", bold)?;

    worksheet.write_with_format(0, 1, "Sales", bold)?;
    worksheet.write(1, 1, 1000)?;
    worksheet.write(2, 1, 1200)?;
    worksheet.write(3, 1, 900)?;
    worksheet.write(4, 1, 1200)?;
    worksheet.write_formula_with_format(5, 1, "=SUBTOTAL(9,B2:B5)", bold)?;

    worksheet.write(6, 0, "South 1")?;
    worksheet.write(7, 0, "South 2")?;
    worksheet.write(8, 0, "South 3")?;
    worksheet.write(9, 0, "South 4")?;
    worksheet.write_with_format(10, 0, "South Total", bold)?;

    worksheet.write(6, 1, 400)?;
    worksheet.write(7, 1, 600)?;
    worksheet.write(8, 1, 500)?;
    worksheet.write(9, 1, 600)?;
    worksheet.write_formula_with_format(10, 1, "=SUBTOTAL(9,B7:B10)", bold)?;

    worksheet.write_with_format(11, 0, "Grand Total", bold)?;
    worksheet.write_formula_with_format(11, 1, "=SUBTOTAL(9,B2:B11)", bold)?;

    // Autofit the columns for clarity.
    worksheet.autofit();

    Ok(())
}
