// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of adding a data validation to a worksheet cell. This validation
//! restricts input to date values in a fixed range.

use rust_xlsxwriter::{DataValidation, DataValidationRule, ExcelDateTime, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write(1, 0, "Enter value in cell D2:")?;

    let data_validation = DataValidation::new().allow_date(DataValidationRule::Between(
        ExcelDateTime::parse_from_str("2025-01-01")?,
        ExcelDateTime::parse_from_str("2025-12-12")?,
    ));

    worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;

    // Save the file.
    workbook.save("data_validation.xlsx")?;

    Ok(())
}
