// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing formatted dates in an Excel
//! worksheet.

use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create some formats to use with the datetimes below.
    let format1 = Format::new().set_num_format("dd/mm/yyyy");
    let format2 = Format::new().set_num_format("mm/dd/yyyy");
    let format3 = Format::new().set_num_format("ddd dd mmm yyyy");
    let format4 = Format::new().set_num_format("dddd, mmmm dd, yyyy");

    // Set the column width for clarity.
    worksheet.set_column_width(0, 30)?;

    // Create a datetime object.
    let datetime = ExcelDateTime::from_ymd(2023, 1, 25)?;

    // Write the datetime with different Excel formats.
    worksheet.write_with_format(0, 0, &datetime, &format1)?;
    worksheet.write_with_format(1, 0, &datetime, &format2)?;
    worksheet.write_with_format(2, 0, &datetime, &format3)?;
    worksheet.write_with_format(3, 0, &datetime, &format4)?;

    workbook.save("datetime.xlsx")?;

    Ok(())
}
