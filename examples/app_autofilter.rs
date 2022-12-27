// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! An example of how to create autofilters with the rust_xlsxwriter library..
//!
//! An autofilter is a way of adding drop down lists to the headers of a 2D
//! range of worksheet data. This allows users to filter the data based on
//! simple criteria so that some data is shown and some is hidden.
//!
//! Note, adding filter criteria isn't currently supported. That will be added
//! in an upcoming version.
//!

use rust_xlsxwriter::{Format, Workbook, Worksheet, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet1 = workbook.add_worksheet();

    // Add some sample filter data.
    populate_autofilter_data(worksheet1)?;

    // Set the autofilter.
    worksheet1.autofilter(0, 0, 50, 3)?;

    // Save the file to disk.
    workbook.save("autofilter.xlsx")?;

    Ok(())
}

// Generate worksheet data to filter on.
pub fn populate_autofilter_data(worksheet: &mut Worksheet) -> Result<(), XlsxError> {
    // The sample data to add to the worksheet.
    let data = vec![
        ("East", "Apple", 9000, "July"),
        ("East", "Apple", 5000, "July"),
        ("South", "Orange", 9000, "September"),
        ("North", "Apple", 2000, "November"),
        ("West", "Apple", 9000, "November"),
        ("South", "Pear", 7000, "October"),
        ("North", "Pear", 9000, "August"),
        ("West", "Orange", 1000, "December"),
        ("West", "Grape", 1000, "November"),
        ("South", "Pear", 10000, "April"),
        ("West", "Grape", 6000, "January"),
        ("South", "Orange", 3000, "May"),
        ("North", "Apple", 3000, "December"),
        ("South", "Apple", 7000, "February"),
        ("West", "Grape", 1000, "December"),
        ("East", "Grape", 8000, "February"),
        ("South", "Grape", 10000, "June"),
        ("West", "Pear", 7000, "December"),
        ("South", "Apple", 2000, "October"),
        ("East", "Grape", 7000, "December"),
        ("North", "Grape", 6000, "April"),
        ("East", "Pear", 8000, "February"),
        ("North", "Apple", 7000, "August"),
        ("North", "Orange", 7000, "July"),
        ("North", "Apple", 6000, "June"),
        ("South", "Grape", 8000, "September"),
        ("West", "Apple", 3000, "October"),
        ("South", "Orange", 10000, "November"),
        ("West", "Grape", 4000, "July"),
        ("North", "Orange", 5000, "August"),
        ("East", "Orange", 1000, "November"),
        ("East", "Orange", 4000, "October"),
        ("North", "Grape", 5000, "August"),
        ("East", "Apple", 1000, "December"),
        ("South", "Apple", 10000, "March"),
        ("East", "Grape", 7000, "October"),
        ("West", "Grape", 1000, "September"),
        ("East", "Grape", 10000, "October"),
        ("South", "Orange", 8000, "March"),
        ("North", "Apple", 4000, "July"),
        ("South", "Orange", 5000, "July"),
        ("West", "Apple", 4000, "June"),
        ("East", "Apple", 5000, "April"),
        ("North", "Pear", 3000, "August"),
        ("East", "Grape", 9000, "November"),
        ("North", "Orange", 8000, "October"),
        ("East", "Apple", 10000, "June"),
        ("South", "Pear", 1000, "December"),
        ("North", "Grape", 10000, "July"),
        ("East", "Grape", 6000, "February"),
    ];

    // Widen the columns for clarity.
    worksheet.set_column_width(0, 12)?;
    worksheet.set_column_width(1, 12)?;
    worksheet.set_column_width(2, 12)?;
    worksheet.set_column_width(3, 12)?;

    // Write the header titles.
    let header_format = Format::new().set_bold();
    worksheet.write_string(0, 0, "Region", &header_format)?;
    worksheet.write_string(0, 1, "Item", &header_format)?;
    worksheet.write_string(0, 2, "Volume", &header_format)?;
    worksheet.write_string(0, 3, "Month", &header_format)?;

    // Write the other worksheet data.
    for (row, data) in data.iter().enumerate() {
        let row = 1 + row as u32;
        worksheet.write_string_only(row, 0, data.0)?;
        worksheet.write_string_only(row, 1, data.1)?;
        worksheet.write_number_only(row, 2, data.2)?;
        worksheet.write_string_only(row, 3, data.3)?;
    }

    Ok(())
}
