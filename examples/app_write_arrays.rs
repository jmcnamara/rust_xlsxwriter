// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of writing arrays of data using the `rust_xlsxwriter` library.
//! Array in this context means Rust arrays or arrays like data types that
//! implement `IntoIterator`. The array must also contain data types that
//! implement `rust_xlsxwriter`'s `IntoExcelData`.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a format for the headings.
    let heading = Format::new().set_bold().set_font_color("#0000CC");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Some array data to write.
    let numbers = [1, 2, 3, 4, 5];
    let words = ["Hello"; 5];
    let matrix = [
        [10, 11, 12, 13, 14],
        [20, 21, 22, 23, 24],
        [30, 31, 32, 33, 34],
    ];

    // Write the array data as columns.
    worksheet.write_with_format(0, 0, "Column data", &heading)?;
    worksheet.write_column(1, 0, numbers)?;
    worksheet.write_column(1, 1, words)?;

    // Write the array data as rows.
    worksheet.write_with_format(0, 4, "Row data", &heading)?;
    worksheet.write_row(1, 4, numbers)?;
    worksheet.write_row(2, 4, words)?;

    // Write the matrix data as an array or rows and as an array of columns.
    worksheet.write_with_format(7, 4, "Row matrix", &heading)?;
    worksheet.write_row_matrix(8, 4, matrix)?;

    worksheet.write_with_format(7, 0, "Column matrix", &heading)?;
    worksheet.write_column_matrix(8, 0, matrix)?;

    // Save the file to disk.
    workbook.save("arrays.xlsx")?;

    Ok(())
}
