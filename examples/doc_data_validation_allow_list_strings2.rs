// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of adding a data validation to a worksheet cell. This validation
//! restricts users to a selection of values from a dropdown list. This example
//! shows how to pre-populate a default choice.

use rust_xlsxwriter::{DataValidation, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write(1, 0, "Select value in cell D2:")?;

    let data_validation =
        DataValidation::new().allow_list_strings(&["Pass", "Fail", "Incomplete"])?;

    worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;

    // Add a default string to the cell with the data validation
    // to pre-populate a default choice.
    worksheet.write(1, 3, "Pass")?;

    // Save the file.
    workbook.save("data_validation.xlsx")?;

    Ok(())
}
