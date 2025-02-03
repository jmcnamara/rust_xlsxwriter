// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of how to group worksheet columns into outlines with
//! collapsed/hidden rows.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a format to use in headings.
    let bold = Format::new().set_bold();

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();

    let data = [50, 20, 15, 25, 65, 80];
    let headings = ["Month", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Total"];

    worksheet.write_row_with_format(0, 0, headings, &bold)?;
    worksheet.write_row(1, 1, data)?;
    worksheet.write_formula_with_format(1, 7, "=SUBTOTAL(9,B2:G2)", &bold)?;

    // Autofit the columns for clarity.
    worksheet.autofit();

    // Add collapse grouping over the sub-total range.
    worksheet.group_columns_collapsed(1, 6)?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
