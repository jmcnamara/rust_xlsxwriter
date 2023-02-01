// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! A simple, getting started, example of some of the features of the
//! rust_xlsxwriter library.

use rust_xlsxwriter::{Format, FormatUnderline, Workbook, XlsxColor, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create a format to use in the worksheet.
    let link_format = Format::new()
        .set_font_color(XlsxColor::Red)
        .set_underline(FormatUnderline::Single);

    // Add a worksheet to the workbook.
    let worksheet1 = workbook.add_worksheet();

    // Set the column width for clarity.
    worksheet1.set_column_width(0, 26)?;

    // Write some url links.
    worksheet1.write_url(0, 0, "https://www.rust-lang.org")?;
    worksheet1.write_url_with_text(1, 0, "https://www.rust-lang.org", "Learn Rust")?;
    worksheet1.write_url_with_format(2, 0, "https://www.rust-lang.org", &link_format)?;

    // Write some internal links.
    worksheet1.write_url(4, 0, "internal:Sheet1!A1")?;
    worksheet1.write_url(5, 0, "internal:Sheet2!C4")?;

    // Write some external links.
    worksheet1.write_url(7, 0, r"file:///C:\Temp\Book1.xlsx")?;
    worksheet1.write_url(8, 0, r"file:///C:\Temp\Book1.xlsx#Sheet1!C4")?;

    // Add another sheet to link to.
    let worksheet2 = workbook.add_worksheet();
    worksheet2.write_string_only(3, 2, "Here I am")?;
    worksheet2.write_url_with_text(4, 2, "internal:Sheet1!A6", "Go back")?;

    // Save the file to disk.
    workbook.save("hyperlinks.xlsx")?;

    Ok(())
}
