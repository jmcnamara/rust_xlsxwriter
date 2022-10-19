// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the scale of the worksheet to fit
//! a defined number of pages vertically and horizontally. This example shows a
//! common use case which is to fit the printed output to 1 page wide but have
//! the height be as long as necessary.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new("worksheet.xlsx");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set the printed output to fit 1 page wide and as long as necessary.
    worksheet.set_print_fit_to_pages(1, 0);

    workbook.close()?;

    Ok(())
}
