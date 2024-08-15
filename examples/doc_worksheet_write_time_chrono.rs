// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing formatted times in an Excel
//! worksheet.

use chrono::NaiveTime;
use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create some formats to use with the times below.
    let format1 = Format::new().set_num_format("h:mm");
    let format2 = Format::new().set_num_format("hh:mm");
    let format3 = Format::new().set_num_format("hh:mm:ss");
    let format4 = Format::new().set_num_format("hh:mm:ss.000");
    let format5 = Format::new().set_num_format("h:mm AM/PM");

    // Set the column width for clarity.
    worksheet.set_column_width(0, 30)?;

    // Create a time object.
    let time = NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap();

    // Write the time with different Excel formats.
    worksheet.write_time_with_format(0, 0, time, &format1)?;
    worksheet.write_time_with_format(1, 0, time, &format2)?;
    worksheet.write_time_with_format(2, 0, time, &format3)?;
    worksheet.write_time_with_format(3, 0, time, &format4)?;
    worksheet.write_time_with_format(4, 0, time, &format5)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
