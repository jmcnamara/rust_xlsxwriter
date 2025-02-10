// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating a format that is a
//! combination of two formats. This example demonstrates that properties
//! in the primary format take precedence.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    // Add some formats.
    let format1 = Format::new().set_font_color("006100").set_bold();
    let format2 = Format::new().set_font_color("9C0006").set_italic();

    // Create new formats based on a merge of two formats.
    let merged1 = format1.merge(&format2);
    let merged2 = format2.merge(&format1);

    // Write some strings with the formats.
    worksheet.write_with_format(0, 0, "Format 1: green and bold", &format1)?;
    worksheet.write_with_format(1, 0, "Format 2: red and italic", &format2)?;
    worksheet.write_with_format(3, 0, "Merged 2 into 1", &merged1)?;
    worksheet.write_with_format(4, 0, "Merged 1 into 2", &merged2)?;

    // Autofit for clarity.
    worksheet.autofit();

    // Save the file.
    workbook.save("formats.xlsx")?;

    Ok(())
}
