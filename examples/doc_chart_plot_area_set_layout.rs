// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of setting the layout of a chart element, in this case the chart
//! plot area.

use rust_xlsxwriter::{Chart, ChartLayout, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 10)?;
    worksheet.write(1, 0, 40)?;
    worksheet.write(2, 0, 50)?;
    worksheet.write(3, 0, 20)?;
    worksheet.write(4, 0, 10)?;
    worksheet.write(5, 0, 50)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);
    chart.add_series().set_values(("Sheet1", 0, 0, 5, 0));

    // Set a chart title and turn off legend for clarity.
    chart.title().set_name("Standard layout");
    chart.legend().set_hidden();

    // Add the stand chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Create a new layout.
    let layout = ChartLayout::new()
        .set_offset(0.20, 0.30)
        .set_dimensions(0.70, 0.50);

    // Apply the layout to the chart plot area.
    chart.plot_area().set_layout(&layout);

    // Set a chart title and turn off legend for clarity.
    chart.title().set_name("Modified layout");
    chart.legend().set_hidden();

    // Add the modified chart to the worksheet.
    worksheet.insert_chart(16, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
