// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the calculation mode of a
//! workbook to manual, so that Excel displays the stored formula result when
//! the file is opened instead of recalculating it.

use rust_xlsxwriter::{CalculationMode, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    // Write a formula and its result.
    worksheet.write_formula(0, 0, "=SUM(A2:A3)")?;
    worksheet.set_formula_result(0, 0, "30");
    worksheet.write_number(1, 0, 10)?;
    worksheet.write_number(2, 0, 20)?;

    // Only recalculate formulas when the user requests it.
    workbook.set_calculation_mode(CalculationMode::Manual);

    workbook.save("workbook.xlsx")?;

    Ok(())
}
