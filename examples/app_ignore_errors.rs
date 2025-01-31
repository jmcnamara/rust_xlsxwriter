// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of turning off worksheet cells errors/warnings using using the
//! `rust_xlsxwriter` library.

use rust_xlsxwriter::{Format, IgnoreError, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a format to use in descriptions.gs
    let bold = Format::new().set_bold();

    // Make the column wider for clarity.
    worksheet.set_column_width(1, 16)?;

    // Write some descriptions for the cells.
    worksheet.write_with_format(1, 1, "Warning:", &bold)?;
    worksheet.write_with_format(2, 1, "Warning turned off:", &bold)?;
    worksheet.write_with_format(4, 1, "Warning:", &bold)?;
    worksheet.write_with_format(5, 1, "Warning turned off:", &bold)?;

    // Write strings that looks like numbers. This will cause an Excel warning.
    worksheet.write_string(1, 2, "123")?;
    worksheet.write_string(2, 2, "123")?;

    // Write a divide by zero formula. This will also cause an Excel warning.
    worksheet.write_formula(4, 2, "=1/0")?;
    worksheet.write_formula(5, 2, "=1/0")?;

    // Turn off some of the warnings:
    worksheet.ignore_error(2, 2, IgnoreError::NumberStoredAsText)?;
    worksheet.ignore_error(5, 2, IgnoreError::FormulaError)?;

    // Save the file to disk.
    workbook.save("ignore_errors.xlsx")?;

    Ok(())
}
