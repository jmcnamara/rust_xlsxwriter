// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing formatted datetimes in an Excel
//! worksheet.

use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create some formats to use with the datetimes below.
    let format1 = Format::new().set_num_format("dd/mm/yyyy hh:mm");
    let format2 = Format::new().set_num_format("mm/dd/yyyy hh:mm");
    let format3 = Format::new().set_num_format("yyyy-mm-ddThh:mm:ss");
    let format4 = Format::new().set_num_format("ddd dd mmm yyyy hh:mm");
    let format5 = Format::new().set_num_format("dddd, mmmm dd, yyyy hh:mm");

    // Set the column width for clarity.
    worksheet.set_column_width(0, 30)?;

    // Create a datetime object.
    let datetime = ExcelDateTime::from_ymd(2023, 1, 25)?.and_hms(12, 30, 0)?;

    // Write the datetime with different Excel formats.
    worksheet.write_datetime_with_format(0, 0, &datetime, &format1)?;
    worksheet.write_datetime_with_format(1, 0, &datetime, &format2)?;
    worksheet.write_datetime_with_format(2, 0, &datetime, &format3)?;
    worksheet.write_datetime_with_format(3, 0, &datetime, &format4)?;
    worksheet.write_datetime_with_format(4, 0, &datetime, &format5)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
