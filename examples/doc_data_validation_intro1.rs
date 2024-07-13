// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of adding a data validation to a worksheet cell. This validation
//! uses an input message to explain to the user what type of input is required.

use rust_xlsxwriter::{DataValidation, DataValidationRule, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write(1, 0, "Enter rating in cell D2:")?;

    let data_validation = DataValidation::new()
        .allow_whole_number(DataValidationRule::Between(1, 5))
        .set_input_title("Enter a star rating!")?
        .set_input_message("Enter rating 1-5.\nWhole numbers only.")?
        .set_error_title("Value outside allowed range")?
        .set_error_message("The input value must be an integer in the range 1-5.")?;

    worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;

    // Save the file.
    workbook.save("data_validation.xlsx")?;

    Ok(())
}
