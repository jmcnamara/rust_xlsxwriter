// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! Example of creating a pivot table with some of the available features.

use rust_xlsxwriter::{
    PivotTable, PivotTableFunction, PivotTableLayout, PivotTableStyle, PivotTableValueField,
    Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet for the source data.
    let worksheet = workbook.add_worksheet().set_name("Data")?;

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

    // Create a pivot table that summarizes the "Volume" by "Region" and "Item",
    // with a filter on "Month".
    let pivot_table = PivotTable::new()
        .set_data_source("Data!A1:D9")
        .add_row_field("Region")
        .add_column_field("Item")
        .add_filter_field("Month")
        .add_value_field(
            &PivotTableValueField::new("Volume")
                .set_function(PivotTableFunction::Sum)
                .set_custom_name("Total Volume")
                .set_num_format("#,##0"),
        )
        .set_style(PivotTableStyle::Medium4)
        .set_layout(PivotTableLayout::Tabular);

    // Add a worksheet for the pivot table and add the pivot table to it.
    let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    worksheet.add_pivot_table(2, 0, &pivot_table)?;

    // Save the file to disk.
    workbook.save("pivot_table.xlsx")?;

    Ok(())
}
