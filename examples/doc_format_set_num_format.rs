// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting different types of Excel number
//! formatting.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new().set_num_format("0.00");
    let format2 = Format::new().set_num_format("0.000");
    let format3 = Format::new().set_num_format("#,##0");
    let format4 = Format::new().set_num_format("#,##0.00");
    let format5 = Format::new().set_num_format("mm/dd/yy");
    let format6 = Format::new().set_num_format("mmm d yyyy");
    let format7 = Format::new().set_num_format("d mmmm yyyy");
    let format8 = Format::new().set_num_format("dd/mm/yyyy hh:mm AM/PM");

    worksheet.write_number(0, 0, 3.1415926, &format1)?;
    worksheet.write_number(1, 0, 3.1415926, &format2)?;
    worksheet.write_number(2, 0, 1234.56, &format3)?;
    worksheet.write_number(3, 0, 1234.56, &format4)?;
    worksheet.write_number(4, 0, 44927.521, &format5)?;
    worksheet.write_number(5, 0, 44927.521, &format6)?;
    worksheet.write_number(6, 0, 44927.521, &format7)?;
    worksheet.write_number(7, 0, 44927.521, &format8)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
