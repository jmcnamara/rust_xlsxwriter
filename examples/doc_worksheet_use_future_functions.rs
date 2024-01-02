// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates different ways to handle writing Future
//! Functions to a worksheet.

use rust_xlsxwriter::{Formula, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // The following is a "Future" function and will generate a "#NAME?" warning
    // in Excel.
    worksheet.write_formula(0, 0, "=ISFORMULA($B$1)")?;

    // The following adds the required prefix. This will work without a warning.
    worksheet.write_formula(1, 0, "=_xlfn.ISFORMULA($B$1)")?;

    // The following uses a Formula object and expands out any future functions.
    // This also works without a warning.
    worksheet.write_formula(
        2,
        0,
        Formula::new("=ISFORMULA($B$1)").use_future_functions(),
    )?;

    // The following expands out all future functions used in the worksheet from
    // this point forward. This also works without a warning.
    worksheet.use_future_functions(true);
    worksheet.write_formula(3, 0, "=ISFORMULA($B$1)")?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
