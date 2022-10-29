// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the indentation level for cell
//! text.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new().set_indent(1);
    let format2 = Format::new().set_indent(2);

    worksheet.write_string_only(0, 0, "Indent 0")?;
    worksheet.write_string(1, 0, "Indent 1", &format1)?;
    worksheet.write_string(2, 0, "Indent 2", &format2)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
