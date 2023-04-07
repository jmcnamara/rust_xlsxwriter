// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of using simulated autofit to automatically adjust the width of
//! worksheet columns based on the data in the cells.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some worksheet data to demonstrate autofitting.
    worksheet.write_string(0, 0, "Foo")?;
    worksheet.write_string(1, 0, "Food")?;
    worksheet.write_string(2, 0, "Foody")?;
    worksheet.write_string(3, 0, "Froody")?;

    worksheet.write_number(0, 1, 12345)?;
    worksheet.write_number(1, 1, 12_345_678)?;
    worksheet.write_number(2, 1, 12345)?;

    worksheet.write_string(0, 2, "Some longer text")?;

    worksheet.write_url(0, 3, "http://ww.google.com")?;
    worksheet.write_url(1, 3, "https://github.com")?;

    // Autofit the worksheet.
    worksheet.autofit();

    // Save the file to disk.
    workbook.save("autofit.xlsx")?;

    Ok(())
}
