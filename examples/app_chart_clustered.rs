// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! A demo of a clustered category chart using rust_xlsxwriter.

use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Types", &bold)?;
    worksheet.write_with_format(0, 1, "Sub Type", &bold)?;
    worksheet.write_with_format(0, 2, "Value 1", &bold)?;
    worksheet.write_with_format(0, 3, "Value 2", &bold)?;
    worksheet.write_with_format(0, 4, "Value 3", &bold)?;

    worksheet.write(1, 0, "Type 1")?;
    worksheet.write(1, 1, "Sub Type A")?;
    worksheet.write(2, 1, "Sub Type B")?;
    worksheet.write(3, 1, "Sub Type C")?;

    worksheet.write(4, 0, "Type 2")?;
    worksheet.write(4, 1, "Sub Type D")?;
    worksheet.write(5, 1, "Sub Type E")?;

    worksheet.write(1, 2, 5000)?;
    worksheet.write(2, 2, 2000)?;
    worksheet.write(3, 2, 250)?;
    worksheet.write(4, 2, 6000)?;
    worksheet.write(5, 2, 500)?;

    worksheet.write(1, 3, 8000)?;
    worksheet.write(2, 3, 3000)?;
    worksheet.write(3, 3, 1000)?;
    worksheet.write(4, 4, 6500)?;
    worksheet.write(5, 3, 300)?;

    worksheet.write(1, 4, 6000)?;
    worksheet.write(2, 4, 4000)?;
    worksheet.write(3, 4, 2000)?;
    worksheet.write(4, 3, 6000)?;
    worksheet.write(5, 4, 200)?;

    // Create a new chart object.
    let mut chart = Chart::new(ChartType::Column);

    // Configure the series. Note, that the categories are 2D ranges (from
    // column A to column B). This creates the clusters.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$B$6")
        .set_values("Sheet1!$C$2:$C$6");

    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$B$6")
        .set_values("Sheet1!$D$2:$D$6");

    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$B$6")
        .set_values("Sheet1!$E$2:$E$6");

    // Set the Excel chart style.
    chart.set_style(37);

    // Turn off the legend.
    chart.legend().set_hidden();

    // Insert the chart into the worksheet.
    worksheet.insert_chart(2, 6, &chart)?;

    workbook.save("chart_clustered.xlsx")?;

    Ok(())
}
