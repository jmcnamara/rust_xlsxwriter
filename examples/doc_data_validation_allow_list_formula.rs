// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of adding a data validation to a worksheet cell. This validation
//! restricts users to a selection of values from a dropdown list. The list data
//! is provided from a cell range.

use rust_xlsxwriter::{DataValidation, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write(1, 0, "Select value in cell D2:")?;

    // Write the string list data to some cells.
    let string_list = ["Pass", "Fail", "Incomplete"];
    worksheet.write_column(1, 5, string_list)?;

    let data_validation = DataValidation::new().allow_list_formula("F2:F4".into());

    worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;

    // Save the file.
    workbook.save("data_validation.xlsx")?;

    Ok(())
}
