// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! A chart example demonstrating deleting/hiding a series name from the chart
//! legend.

use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 30)?;
    worksheet.write(1, 0, 20)?;
    worksheet.write(2, 0, 40)?;
    worksheet.write(0, 1, 10)?;
    worksheet.write(1, 1, 10)?;
    worksheet.write(2, 1, 10)?;
    worksheet.write(0, 2, 20)?;
    worksheet.write(1, 2, 15)?;
    worksheet.write(2, 2, 30)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);

    // Add a series whose name will appear in the legend.
    chart.add_series().set_values("Sheet1!$A$1:$A$3");

    // Add a series but delete/hide its names from the legend.
    chart
        .add_series()
        .set_values("Sheet1!$B$1:$B$3")
        .delete_from_legend(true);

    // Add a series whose name will appear in the legend.
    chart.add_series().set_values("Sheet1!$C$1:$C$3");

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 3, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
