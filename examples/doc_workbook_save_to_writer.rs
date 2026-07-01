// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating a simple workbook to some types
//! that implement the `Write` trait like a file and a buffer.

use std::fs::File;
use std::io::Write;

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();
    worksheet.write_string(0, 0, "Hello")?;

    // 1. Save the file to a File object.
    let file = File::create("workbook1.xlsx")?;
    workbook.save_to_writer(file)?;

    // 2. Save the file to a buffer.
    let mut buf = Vec::new();
    workbook.save_to_writer(&mut buf)?;

    // Write the buffer to a file for the sake of the example.
    let mut file = File::create("workbook2.xlsx")?;
    Write::write_all(&mut file, &buf)?;

    Ok(())
}
