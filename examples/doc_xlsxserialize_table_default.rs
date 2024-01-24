// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of serializing Serde derived structs to an Excel worksheet using
//! `rust_xlsxwriter` and the `XlsxSerialize` trait.

use rust_xlsxwriter::{Workbook, XlsxError, XlsxSerialize};
use serde::Serialize;

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a serializable struct.
    #[derive(XlsxSerialize, Serialize)]
    #[xlsx(table_default)]
    struct Produce {
        fruit: &'static str,
        cost: f64,
    }

    // Create some data instances.
    let items = [
        Produce {
            fruit: "Peach",
            cost: 1.05,
        },
        Produce {
            fruit: "Plum",
            cost: 0.15,
        },
        Produce {
            fruit: "Pear",
            cost: 0.75,
        },
    ];

    // Set the serialization location and headers.
    worksheet.set_serialize_headers::<Produce>(0, 0)?;

    // Serialize the data.
    worksheet.serialize(&items)?;

    // Save the file to disk.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
