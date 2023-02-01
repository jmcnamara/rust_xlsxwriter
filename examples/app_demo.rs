// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! A simple, getting started, example of some of the features of the
//! rust_xlsxwriter library.

use chrono::NaiveDate;
use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create some formats to use in the worksheet.
    let bold_format = Format::new().set_bold();
    let decimal_format = Format::new().set_num_format("0.000");
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let merge_format = Format::new()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center);

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set the column width for clarity.
    worksheet.set_column_width(0, 22)?;

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
    let date = NaiveDate::from_ymd_opt(2023, 1, 25).unwrap();
    worksheet.write_date(6, 0, &date, &date_format)?;

    // Write some links.
    worksheet.write_url(7, 0, "https://www.rust-lang.org")?;
    worksheet.write_url_with_text(8, 0, "https://www.rust-lang.org", "Learn Rust")?;

    // Write some merged cells.
    worksheet.merge_range(9, 0, 9, 1, "Merged cells", &merge_format)?;

    // Insert an image.
    let image = Image::new("examples/rust_logo.png")?;
    worksheet.insert_image(1, 2, &image)?;

    // Save the file to disk.
    workbook.save("demo.xlsx")?;

    Ok(())
}
