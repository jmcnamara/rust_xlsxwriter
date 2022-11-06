// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! A simple, getting started, example of some of the features of the
//! rust_xlsxwriter library.

use chrono::NaiveDate;
use rust_xlsxwriter::{Format, Workbook, XlsxAlign, XlsxBorder, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create some formats to use in the worksheet.
    let bold_format = Format::new().set_bold();
    let decimal_format = Format::new().set_num_format("0.000");
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let merge_format = Format::new()
        .set_border(XlsxBorder::Thin)
        .set_align(XlsxAlign::Center);

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set the column width for clarity.
    worksheet.set_column_width(0, 15)?;

    // Write a string without formatting.
    worksheet.write_string_only(0, 0, "Hello")?;

    // Write a string with the bold format defined above.
    worksheet.write_string(1, 0, "World", &bold_format)?;

    // Write some numbers.
    worksheet.write_number_only(2, 0, 1)?;
    worksheet.write_number_only(3, 0, 2.34)?;

    // Write a number with formatting.
    worksheet.write_number(4, 0, 3.00, &decimal_format)?;

    // Write a formula.
    worksheet.write_formula_only(5, 0, "=SIN(PI()/4)")?;

    // Write a date.
    let date = NaiveDate::from_ymd(2023, 1, 25);
    worksheet.write_date(6, 0, date, &date_format)?;

    // Write some merged cells.
    worksheet.merge_range(7, 0, 7, 1, "Merged cells", &merge_format)?;

    // Save the file to disk.
    workbook.save("demo.xlsx")?;

    Ok(())
}
