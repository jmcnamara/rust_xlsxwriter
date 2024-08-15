// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing datetimes that take an implicit
//! format from the column formatting.

use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create some formats to use with the datetimes below.
    let format1 = Format::new().set_num_format("dd/mm/yyyy hh:mm");
    let format2 = Format::new().set_num_format("mm/dd/yyyy hh:mm");
    let format3 = Format::new().set_num_format("yyyy-mm-ddThh:mm:ss");

    // Set the column formats.
    worksheet.set_column_format(0, &format1)?;
    worksheet.set_column_format(1, &format2)?;
    worksheet.set_column_format(2, &format3)?;

    // Set the column widths for clarity.
    worksheet.set_column_width(0, 20)?;
    worksheet.set_column_width(1, 20)?;
    worksheet.set_column_width(2, 20)?;

    // Create a datetime object.
    let datetime = ExcelDateTime::from_ymd(2023, 1, 25)?.and_hms(12, 30, 0)?;

    // Write the datetime without a formats. The dates will get the column
    // format instead.
    worksheet.write_datetime(0, 0, &datetime)?;
    worksheet.write_datetime(0, 1, &datetime)?;
    worksheet.write_datetime(0, 2, &datetime)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
