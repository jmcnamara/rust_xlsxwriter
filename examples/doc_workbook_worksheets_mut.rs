// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates operating on the vector of all the
//! worksheets in a workbook.

use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add three worksheets to the workbook.
    let _ = workbook.add_worksheet();
    let _ = workbook.add_worksheet();
    let _ = workbook.add_worksheet();

    let worksheet4 = Worksheet::new();

    workbook.worksheets_mut().push(worksheet4);

    // Write the same data to all three worksheets.
    for worksheet in workbook.worksheets_mut() {
        worksheet.write_string_only(0, 0, "Hello")?;
        worksheet.write_number_only(1, 0, 12345)?;
    }

    // If you are careful you can use standard slice operations.
    workbook.worksheets_mut().swap(0, 1);

    workbook.save("workbook.xlsx")?;

    Ok(())
}
