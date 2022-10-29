// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing an array formulas to a worksheet.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some test data.
    worksheet.write_number_only(0, 1, 500)?;
    worksheet.write_number_only(0, 2, 300)?;
    worksheet.write_number_only(1, 1, 10)?;
    worksheet.write_number_only(1, 2, 15)?;

    // Write an array formula that returns a single value.
    worksheet.write_array_formula_only(0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}")?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
