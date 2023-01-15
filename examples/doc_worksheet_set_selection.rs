// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates selecting cells in worksheets. The order
//! of selection within the range depends on the order of `first` and `last`.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet1 = workbook.add_worksheet();
    worksheet1.set_selection(3, 2, 3, 2)?; // Cell C4

    let worksheet2 = workbook.add_worksheet();
    worksheet2.set_selection(3, 2, 6, 6)?; // Cells C4 to G7.

    let worksheet3 = workbook.add_worksheet();
    worksheet3.set_selection(6, 6, 3, 2)?; // Cells G7 to C4.

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
