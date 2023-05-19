// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates manually setting the result of a formula.
//! Note, this is only required for non-Excel applications that don't calculate
//! formula results.

use rust_xlsxwriter::{Formula, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Using the formula string syntax.
    worksheet
        .write_formula(0, 0, "1+1")?
        .set_formula_result(0, 0, "2");

    // Or using a Formula type.
    worksheet.write_formula(1, 0, Formula::new("2+2").set_result("4"))?;

    workbook.save("formulas.xlsx")?;

    Ok(())
}
