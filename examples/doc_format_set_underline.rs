// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting underline properties for a
//! format.

use rust_xlsxwriter::{Format, FormatUnderline, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new().set_underline(FormatUnderline::None);
    let format2 = Format::new().set_underline(FormatUnderline::Single);
    let format3 = Format::new().set_underline(FormatUnderline::Double);
    let format4 = Format::new().set_underline(FormatUnderline::SingleAccounting);
    let format5 = Format::new().set_underline(FormatUnderline::DoubleAccounting);

    worksheet.write_string_with_format(0, 0, "None", &format1)?;
    worksheet.write_string_with_format(1, 0, "Single", &format2)?;
    worksheet.write_string_with_format(2, 0, "Double", &format3)?;
    worksheet.write_string_with_format(3, 0, "Single Accounting", &format4)?;
    worksheet.write_string_with_format(4, 0, "Double Accounting", &format5)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
