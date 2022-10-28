// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! An example of creating a simple Excel xlsx file in an in memory Vec<u8>
//! buffer using the rust_xlsxwriter library.
//!
//! Note, although the final file will be in memory the intermediate files that
//! make up the xlsx file are written to temporary files on disk before the
//! final file is assembled.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file.
    let mut workbook = Workbook::new_from_buffer();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write a string to cell (0, 0) = A1.
    worksheet.write_string_only(0, 0, "Hello")?;

    // Write a number to cell (1, 0) = A2.
    worksheet.write_number_only(1, 0, 12345)?;

    // Get the file data in a Vec<u8> buffer.
    let buf = workbook.close_to_buffer()?;

    // For the sake of this example we write the data to a file.
    let file = std::fs::File::create("from_buffer.xlsx")?;
    let mut writer = std::io::BufWriter::new(file);
    std::io::Write::write_all(&mut writer, &buf)?;

    Ok(())
}
