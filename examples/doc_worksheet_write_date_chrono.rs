// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing formatted dates in an Excel
//! worksheet.

use chrono::NaiveDate;
use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create some formats to use with the dates below.
    let format1 = Format::new().set_num_format("dd/mm/yyyy");
    let format2 = Format::new().set_num_format("mm/dd/yyyy");
    let format3 = Format::new().set_num_format("yyyy-mm-dd");
    let format4 = Format::new().set_num_format("ddd dd mmm yyyy");
    let format5 = Format::new().set_num_format("dddd, mmmm dd, yyyy");

    // Set the column width for clarity.
    worksheet.set_column_width(0, 30)?;

    // Create a date object.
    let date = NaiveDate::from_ymd_opt(2023, 1, 25).unwrap();

    // Write the date with different Excel formats.
    worksheet.write_date_with_format(0, 0, &date, &format1)?;
    worksheet.write_date_with_format(1, 0, &date, &format2)?;
    worksheet.write_date_with_format(2, 0, &date, &format3)?;
    worksheet.write_date_with_format(3, 0, &date, &format4)?;
    worksheet.write_date_with_format(4, 0, &date, &format5)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
