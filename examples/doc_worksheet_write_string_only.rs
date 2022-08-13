// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing some UTF-8 strings to a
//! worksheet. The UTF-8 encoding is the only encoding supported by the Excel
//! file format.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file.
    let mut workbook = Workbook::new("strings.xlsx");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some strings to the worksheet.
    worksheet.write_string_only(0, 0, "السلام عليكم")?;
    worksheet.write_string_only(1, 0, "Dobrý den")?;
    worksheet.write_string_only(2, 0, "Hello")?;
    worksheet.write_string_only(3, 0, "שָׁלוֹם")?;
    worksheet.write_string_only(4, 0, "नमस्ते")?;
    worksheet.write_string_only(5, 0, "こんにちは")?;
    worksheet.write_string_only(6, 0, "안녕하세요")?;
    worksheet.write_string_only(7, 0, "你好")?;
    worksheet.write_string_only(8, 0, "Olá")?;
    worksheet.write_string_only(9, 0, "Здравствуйте")?;
    worksheet.write_string_only(10, 0, "Hola")?;

    workbook.close()?;

    Ok(())
}
