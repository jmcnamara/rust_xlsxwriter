// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing formatted datetimes parsed from
//! strings.

use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create some formats to use with the datetimes below.
    let format1 = Format::new().set_num_format("hh:mm:ss");
    let format2 = Format::new().set_num_format("yyyy-mm-dd");
    let format3 = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");

    // Set the column width for clarity.
    worksheet.set_column_width(0, 30)?;

    // Create different datetime objects.
    let datetime1 = ExcelDateTime::parse_from_str("12:30")?;
    let datetime2 = ExcelDateTime::parse_from_str("12:30:45")?;
    let datetime3 = ExcelDateTime::parse_from_str("12:30:45.5")?;
    let datetime4 = ExcelDateTime::parse_from_str("2023-01-31")?;
    let datetime5 = ExcelDateTime::parse_from_str("2023-01-31 12:30:45")?;
    let datetime6 = ExcelDateTime::parse_from_str("2023-01-31T12:30:45Z")?;

    // Write the datetime with different Excel formats.
    worksheet.write_with_format(0, 0, &datetime1, &format1)?;
    worksheet.write_with_format(1, 0, &datetime2, &format1)?;
    worksheet.write_with_format(2, 0, &datetime3, &format1)?;
    worksheet.write_with_format(3, 0, &datetime4, &format2)?;
    worksheet.write_with_format(4, 0, &datetime5, &format3)?;
    worksheet.write_with_format(5, 0, &datetime6, &format3)?;

    workbook.save("datetime.xlsx")?;

    Ok(())
}
