// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of adding a data validation to a worksheet cell. This validation
//! restricts input to email addresses based on a custom formula that checks
//! (too simply) for a `@` in the string.

use rust_xlsxwriter::{DataValidation, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write(1, 0, "Enter valid email in C2:")?;

    let data_validation = DataValidation::new().allow_custom(r#"=ISNUMBER(FIND("@",C2))"#.into());

    worksheet.add_data_validation(1, 2, 1, 2, &data_validation)?;

    // Save the file.
    workbook.save("data_validation.xlsx")?;

    Ok(())
}
