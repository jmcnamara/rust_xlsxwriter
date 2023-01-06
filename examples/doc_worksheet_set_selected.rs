// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates selecting worksheet in a workbook. The
//! active worksheet is selected by default so in this example the first two
//! worksheets are selected.

use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet1 = Worksheet::new();
    let worksheet3 = Worksheet::new();
    let mut worksheet2 = Worksheet::new();

    worksheet2.set_selected(true);

    workbook.push_worksheet(worksheet1);
    workbook.push_worksheet(worksheet2);
    workbook.push_worksheet(worksheet3);

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
