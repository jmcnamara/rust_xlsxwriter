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

    // Create a formats to use with the datetimes below.
    let format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");

    // Set the column width for clarity.
    worksheet.set_column_width(0, 30)?;

    // Create a datetime object.
    let datetime1 = ExcelDateTime::from_serial_datetime(1.5)?;
    let datetime2 = ExcelDateTime::from_serial_datetime(36526.61)?;
    let datetime3 = ExcelDateTime::from_serial_datetime(44951.72)?;

    // Write the formatted datetime.
    worksheet.write_with_format(0, 0, &datetime1, &format)?;
    worksheet.write_with_format(1, 0, &datetime2, &format)?;
    worksheet.write_with_format(2, 0, &datetime3, &format)?;

    workbook.save("datetime.xlsx")?;

    Ok(())
}
