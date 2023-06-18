// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing formatted datetimes in an Excel
//! worksheet.

use rust_xlsxwriter::{ExcelDateTime, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set the column width for clarity.
    worksheet.set_column_width(0, 30)?;

    // Create a datetime object.
    let datetime1 = ExcelDateTime::from_serial_datetime(1.5)?;
    let datetime2 = ExcelDateTime::from_serial_datetime(36526.61)?;
    let datetime3 = ExcelDateTime::from_serial_datetime(44951.72)?;

    // Write the datetime with the default "yyyy\\-mm\\-dd\\ hh:mm:ss" format.
    worksheet.write(0, 0, &datetime1)?;
    worksheet.write(1, 0, &datetime2)?;
    worksheet.write(2, 0, &datetime3)?;

    workbook.save("datetime.xlsx")?;

    Ok(())
}
