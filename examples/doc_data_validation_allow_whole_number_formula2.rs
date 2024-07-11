// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of adding a data validation to a worksheet cell. This validation
//! restricts input to integer values based on a value from another cell.

use rust_xlsxwriter::{DataValidation, DataValidationRule, Formula, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write(0, 0, "Upper limit:")?;
    worksheet.write(0, 3, 10)?;
    worksheet.write(1, 0, "Enter value in cell D2:")?;

    let data_validation = DataValidation::new()
        .allow_whole_number_formula(DataValidationRule::LessThanOrEqualTo(Formula::new("=D1")));

    worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;

    // Save the file.
    workbook.save("data_validation.xlsx")?;

    Ok(())
}
