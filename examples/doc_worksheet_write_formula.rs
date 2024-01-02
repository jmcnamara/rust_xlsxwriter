// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing formulas with formatting to a
//! worksheet.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some formulas to the worksheet.
    worksheet.write_formula(0, 0, "=B3 + B4")?;
    worksheet.write_formula(1, 0, "=SIN(PI()/4)")?;
    worksheet.write_formula(2, 0, "=SUM(B1:B5)")?;
    worksheet.write_formula(3, 0, r#"=IF(A3>1,"Yes", "No")"#)?;
    worksheet.write_formula(4, 0, "=AVERAGE(1, 2, 3, 4)")?;
    worksheet.write_formula(5, 0, r#"=DATEVALUE("1-Jan-2023")"#)?;

    workbook.save("formulas.xlsx")?;

    Ok(())
}
