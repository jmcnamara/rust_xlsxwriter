// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates manually setting the result of a formula.
//! Note, this is only required for non-Excel applications that don't calculate
//! formula results.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    worksheet
        .write_formula_only(0, 0, "1+1")?
        .set_formula_result(0, 0, "2");

    workbook.save("formulas.xlsx")?;

    Ok(())
}
