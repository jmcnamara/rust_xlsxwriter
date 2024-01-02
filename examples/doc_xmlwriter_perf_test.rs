// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Simple performance test to exercise xmlwriter without hitting the
//! worksheet::write_data_table() fast path.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    for _ in 0..1000 {
        let _ = workbook.add_worksheet();
    }

    workbook.save("workbook.xlsx")?;

    Ok(())
}
