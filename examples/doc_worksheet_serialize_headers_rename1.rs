// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates renaming fields during serialization by
//! using Serde field attributes.

use rust_xlsxwriter::{Workbook, XlsxError};
use serde::{Deserialize, Serialize};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a serializable struct. Note the serde attributes.
    #[derive(Deserialize, Serialize)]
    struct Produce {
        #[serde(rename = "Item")]
        fruit: &'static str,

        #[serde(rename = "Price")]
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
    worksheet.deserialize_headers::<Produce>(0, 0)?;

    // Serialize the data.
    worksheet.serialize(&items)?;

    // Save the file.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
