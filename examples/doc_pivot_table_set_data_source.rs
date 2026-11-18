// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! Example of creating a pivot table and setting the data source.

use rust_xlsxwriter::{PivotTable, PivotTableValueField, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet for the source data.
    let worksheet = workbook.add_worksheet();

    // Some sample data to summarize.
    let data = [
        ("East", "Apple", 9000, "July"),
        ("West", "Pear", 5000, "July"),
        ("North", "Apple", 3000, "August"),
        ("East", "Grape", 8000, "August"),
        ("West", "Apple", 6000, "September"),
        ("North", "Pear", 4000, "September"),
        ("East", "Pear", 7000, "October"),
        ("South", "Grape", 2000, "October"),
    ];

    // Write the header row and the data.
    worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    for (row, (region, item, volume, month)) in data.iter().enumerate() {
        let row = row as u32 + 1;
        worksheet.write(row, 0, *region)?;
        worksheet.write(row, 1, *item)?;
        worksheet.write(row, 2, *volume)?;
        worksheet.write(row, 3, *month)?;
    }

    // Create a pivot table that sums the "Volume" by "Region". Set the data
    // source to point to the data on the first worksheet.
    let pivot_table = PivotTable::new()
        .set_data_source("Sheet1!A1:D9")
        .add_row_field("Region")
        .add_value_field(&PivotTableValueField::new("Volume"));

    // Add a worksheet for the pivot table and add the pivot table to it.
    let worksheet = workbook.add_worksheet();
    worksheet.add_pivot_table(2, 0, &pivot_table)?;

    // Save the file to disk.
    workbook.save("pivot_table.xlsx")?;

    Ok(())
}
