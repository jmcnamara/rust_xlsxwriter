// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example showing all 48 default chart styles available in Excel 2007
//! using rust_xlsxwriter. Note, these styles are not the same as the styles
//! available in Excel 2013 and later.

use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let chart_types = vec![
        ("Column", ChartType::Column),
        ("Area", ChartType::Area),
        ("Line", ChartType::Line),
        ("Pie", ChartType::Pie),
    ];

    // Create a worksheet with 48 charts in each of the available styles, for
    // each of the chart types above.
    for (name, chart_type) in chart_types {
        let worksheet = workbook.add_worksheet().set_name(name)?.set_zoom(30);
        let mut chart = Chart::new(chart_type);
        chart.add_series().set_values("Data!$A$1:$A$6");
        let mut style = 1;

        for row_num in (0..90).step_by(15) {
            for col_num in (0..64).step_by(8) {
                chart.set_style(style);
                chart.title().set_name(&format!("Style {style}"));
                chart.legend().set_hidden();
                worksheet.insert_chart(row_num as u32, col_num as u16, &chart)?;
                style += 1;
            }
        }
    }

    // Create a worksheet with data for the charts.
    let data_worksheet = workbook.add_worksheet().set_name("Data")?;
    data_worksheet.write(0, 0, 10)?;
    data_worksheet.write(1, 0, 40)?;
    data_worksheet.write(2, 0, 50)?;
    data_worksheet.write(3, 0, 20)?;
    data_worksheet.write(4, 0, 10)?;
    data_worksheet.write(5, 0, 50)?;
    data_worksheet.set_hidden(true);

    workbook.save("chart_styles.xlsx")?;

    Ok(())
}
