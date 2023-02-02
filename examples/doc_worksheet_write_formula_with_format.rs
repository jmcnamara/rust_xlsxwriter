// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing formulas with formatting to a
//! worksheet.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create some formats to use in the worksheet.
    let bold_format = Format::new().set_bold();
    let italic_format = Format::new().set_italic();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some formulas with formatting.
    worksheet.write_formula_with_format(0, 0, "=1+2+3", &bold_format)?;
    worksheet.write_formula_with_format(1, 0, "=A1*2", &bold_format)?;
    worksheet.write_formula_with_format(2, 0, "=SIN(PI()/4)", &italic_format)?;
    worksheet.write_formula_with_format(3, 0, "=AVERAGE(1, 2, 3, 4)", &italic_format)?;

    workbook.save("formulas.xlsx")?;

    Ok(())
}
