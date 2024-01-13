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
    struct Produce {
        fruit: &'static str,
        cost: f64,

        #[serde(skip)]
        #[allow(dead_code)]
        in_stock: bool,
    }

    // Create some data instances.
    let item1 = Produce {
        fruit: "Peach",
        cost: 1.05,
        in_stock: true,
    };

    let item2 = Produce {
        fruit: "Plum",
        cost: 0.15,
        in_stock: true,
    };

    let item3 = Produce {
        fruit: "Pear",
        cost: 0.75,
        in_stock: false,
    };

    // Set the serialization location and headers.
    worksheet.set_serialize_headers::<Produce>(0, 0)?;

    // Serialize the data.
    worksheet.serialize(&item1)?;
    worksheet.serialize(&item2)?;
    worksheet.serialize(&item3)?;

    // Save the file to disk.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
