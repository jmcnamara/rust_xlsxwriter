// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting different types of Excel number
//! formatting.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    // Set column width for clarity.
    worksheet.set_column_width(0, 20)?;

    let format1 = Format::new().set_num_format("0.00");
    let format2 = Format::new().set_num_format("0.000");
    let format3 = Format::new().set_num_format("#,##0");
    let format4 = Format::new().set_num_format("#,##0.00");
    let format5 = Format::new().set_num_format("mm/dd/yy");
    let format6 = Format::new().set_num_format("mmm d yyyy");
    let format7 = Format::new().set_num_format("d mmmm yyyy");
    let format8 = Format::new().set_num_format("dd/mm/yyyy hh:mm AM/PM");

    worksheet.write_number_with_format(0, 0, 1.23456, &format1)?;
    worksheet.write_number_with_format(1, 0, 1.23456, &format2)?;
    worksheet.write_number_with_format(2, 0, 1234.56, &format3)?;
    worksheet.write_number_with_format(3, 0, 1234.56, &format4)?;
    worksheet.write_number_with_format(4, 0, 44927.521, &format5)?;
    worksheet.write_number_with_format(5, 0, 44927.521, &format6)?;
    worksheet.write_number_with_format(6, 0, 44927.521, &format7)?;
    worksheet.write_number_with_format(7, 0, 44927.521, &format8)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
