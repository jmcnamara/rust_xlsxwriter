// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! An example of using a simulated Column Autofit using the rust_xlsxwriter
//! library.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some data to the worksheet.

    // Columns with explicitly defined widths aren't auto-fitted.
    worksheet.set_column_width(0, 18)?;
    worksheet.write_string_only(0, 0, "User set widths")?;
    worksheet.write_string_only(1, 0, "aren't changed")?;

    // Columns without explicitly defined widths are auto-fitted.
    worksheet.write_string_only(0, 1, "Implicit widths")?;
    worksheet.write_string_only(1, 1, "are changed")?;

    // Add some similar words of different lengths.
    worksheet.write_string_only(0, 2, "Fo")?;
    worksheet.write_string_only(1, 2, "Foo")?;
    worksheet.write_string_only(2, 2, "Food")?;
    worksheet.write_string_only(3, 2, "Frood")?;
    worksheet.write_string_only(4, 2, "Froody")?;

    // Add some numbers.
    worksheet.write_number_only(0, 3, 123)?;
    worksheet.write_number_only(0, 4, 1234)?;
    worksheet.write_number_only(0, 5, 12345)?;

    // Run autofit after the data have been added to the worksheet.
    worksheet.set_autofit();

    // Save the file to disk.
    workbook.save("autofit.xlsx")?;

    Ok(())
}
