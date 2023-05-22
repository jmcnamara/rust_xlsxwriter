// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates set the tab color of worksheets.

use rust_xlsxwriter::{Color, Workbook, Worksheet, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let mut worksheet1 = Worksheet::new();
    let mut worksheet2 = Worksheet::new();
    let mut worksheet3 = Worksheet::new();
    let mut worksheet4 = Worksheet::new();

    worksheet1.set_tab_color(Color::Red);
    worksheet2.set_tab_color(Color::Green);
    worksheet3.set_tab_color(Color::RGB(0xFF9900));

    // worksheet4 will have the default color.
    worksheet4.set_active(true);

    workbook.push_worksheet(worksheet1);
    workbook.push_worksheet(worksheet2);
    workbook.push_worksheet(worksheet3);
    workbook.push_worksheet(worksheet4);

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
