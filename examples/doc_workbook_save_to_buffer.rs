// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating a simple workbook to a Vec<u8>
//! buffer.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();
    worksheet.write_string(0, 0, "Hello")?;

    let buf = workbook.save_to_buffer()?;

    println!("File size: {}", buf.len());

    Ok(())
}
