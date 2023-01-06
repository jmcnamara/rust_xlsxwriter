// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! Example of how to use the rust_xlsxwriter to write simple array formulas.

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

    worksheet.write_number_only(4, 1, 1)?;
    worksheet.write_number_only(4, 2, 20234)?;
    worksheet.write_number_only(5, 1, 2)?;
    worksheet.write_number_only(5, 2, 21003)?;
    worksheet.write_number_only(6, 1, 3)?;
    worksheet.write_number_only(6, 2, 10000)?;

    // Write an array formula that returns a single value.
    worksheet.write_array_formula_only(0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}")?;

    // The curly brackets and equal sign are optional.
    worksheet.write_array_formula_only(1, 0, 1, 0, "SUM(B1:C1*B2:C2)")?;

    // Write an array formula that returns a range of values.
    worksheet.write_array_formula_only(4, 0, 6, 0, "{=TREND(C5:C7,B5:B7)}")?;

    // Save the file to disk.
    workbook.save("array_formula.xlsx")?;

    Ok(())
}
