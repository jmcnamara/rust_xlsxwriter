// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! Example of cell locking and formula hiding in an Excel worksheet
//! `rust_xlsxwriter` library.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create some format objects.
    let unlocked = Format::new().set_unlocked();
    let hidden = Format::new().set_hidden();

    // Protect the worksheet to turn on cell locking.
    worksheet.protect();

    // Examples of cell locking and hiding.
    worksheet.write_string(0, 0, "Cell B1 is locked. It cannot be edited.")?;
    worksheet.write_formula(0, 1, "=1+2")?; // Locked by default.

    worksheet.write_string(1, 0, "Cell B2 is unlocked. It can be edited.")?;
    worksheet.write_formula_with_format(1, 1, "=1+2", &unlocked)?;

    worksheet.write_string(2, 0, "Cell B3 is hidden. The formula isn't visible.")?;
    worksheet.write_formula_with_format(2, 1, "=1+2", &hidden)?;

    worksheet.write_string(4, 0, "Use Menu -> Review -> Unprotect Sheet")?;
    worksheet.write_string(5, 0, "to remove the worksheet protection.")?;

    worksheet.autofit();

    // Save the file to disk.
    workbook.save("worksheet_protection.xlsx")?;

    Ok(())
}
