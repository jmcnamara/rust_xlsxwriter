// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates operating on the vector of all the
//! worksheets in a workbook. The non mutable version of this method is less
//! useful than `workbook.worksheets_mut()`.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add three worksheets to the workbook.
    let _ = workbook.add_worksheet();
    let _ = workbook.add_worksheet();
    let _ = workbook.add_worksheet();

    // Get some information from all three worksheets.
    for worksheet in workbook.worksheets() {
        println!("{}", worksheet.name());
    }

    workbook.save("workbook.xlsx")?;

    Ok(())
}
