// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the worksheet paper size/type for
//! the printed output.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set the printer paper size.
    worksheet.set_paper_size(9); // A4 paper size.

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
