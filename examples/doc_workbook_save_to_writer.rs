// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating a simple workbook to some types
//! that implement the `Write` trait like a file and a buffer.

use std::fs::File;
use std::io::{Cursor, Write};

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();
    worksheet.write_string(0, 0, "Hello")?;

    // Save the file to a File object.
    let file = File::create("workbook1.xlsx")?;
    workbook.save_to_writer(file)?;

    // Save the file to a buffer. It is wrapped in a Cursor because it need to
    // implement the `Seek` trait. See also the `workbook.save_to_buffer()`
    // method for an alternative approach.
    let mut cursor = Cursor::new(Vec::new());
    workbook.save_to_writer(&mut cursor)?;

    // Write the buffer to a file for the sake of the example.
    let buf = cursor.into_inner();
    let mut file = File::create("workbook2.xlsx")?;
    Write::write_all(&mut file, &buf)?;

    Ok(())
}
