// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting an implicit (without newline)
//! text wrap and a user defined text wrap (with newlines).

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new().set_text_wrap();

    worksheet.write_string_only(0, 0, "Some text that isn't wrapped")?;
    worksheet.write_string(1, 0, "Some text that is wrapped", &format1)?;
    worksheet.write_string(2, 0, "Some text\nthat is\nwrapped\nat newlines", &format1)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
