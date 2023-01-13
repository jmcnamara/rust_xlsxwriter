// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates operating on the vector of all the
//! worksheets in a workbook.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add three worksheets to the workbook.
    let _worksheet1 = workbook.add_worksheet();
    let _worksheet2 = workbook.add_worksheet();
    let _worksheet3 = workbook.add_worksheet();

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
