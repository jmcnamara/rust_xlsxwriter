// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates handling NaN and Infinity values and also
//! setting custom string representations.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write NaN and Infinity default values.
    worksheet.write(0, 0, "Default:")?;
    worksheet.write(0, 1, f64::NAN)?;
    worksheet.write(1, 1, f64::INFINITY)?;
    worksheet.write(2, 1, f64::NEG_INFINITY)?;

    // Overwrite the default values.
    worksheet.set_nan_value("Nan");
    worksheet.set_infinity_value("Infinity");
    worksheet.set_neg_infinity_value("NegInfinity");

    // Write NaN and Infinity custom values.
    worksheet.write(4, 0, "Custom:")?;
    worksheet.write(4, 1, f64::NAN)?;
    worksheet.write(5, 1, f64::INFINITY)?;
    worksheet.write(6, 1, f64::NEG_INFINITY)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
