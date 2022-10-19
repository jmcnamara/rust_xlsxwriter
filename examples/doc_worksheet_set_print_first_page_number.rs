// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the page number on the printed
//! page.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new("worksheet.xlsx");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    worksheet.set_header("&CPage &P of &N");
    worksheet.set_print_first_page_number(2);

    workbook.close()?;

    Ok(())
}
