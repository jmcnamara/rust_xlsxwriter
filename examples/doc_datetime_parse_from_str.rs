// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing formatted datetimes parsed from
//! strings.

use rust_xlsxwriter::{ExcelDateTime, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set the column width for clarity.
    worksheet.set_column_width(0, 30)?;

    // Create a datetime object.
    let datetime1 = ExcelDateTime::parse_from_str("12:30")?;
    let datetime2 = ExcelDateTime::parse_from_str("12:30:45")?;
    let datetime3 = ExcelDateTime::parse_from_str("12:30:45.5")?;
    let datetime4 = ExcelDateTime::parse_from_str("2023-01-31")?;
    let datetime5 = ExcelDateTime::parse_from_str("2023-01-31 12:30:45")?;
    let datetime6 = ExcelDateTime::parse_from_str("2023-01-31T12:30:45Z")?;

    // Write the dates and times with the default number formats.
    worksheet.write(0, 0, &datetime1)?;
    worksheet.write(1, 0, &datetime2)?;
    worksheet.write(2, 0, &datetime3)?;
    worksheet.write(3, 0, &datetime4)?;
    worksheet.write(4, 0, &datetime5)?;
    worksheet.write(5, 0, &datetime6)?;

    workbook.save("datetime.xlsx")?;

    Ok(())
}
