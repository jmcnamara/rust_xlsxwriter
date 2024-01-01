// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates turning off headers during serialization.
//! The example in columns "D:E" have the headers turned off.
//!
use rust_xlsxwriter::{SerializeHeadersOptions, Workbook, XlsxError};
use serde::{Deserialize, Serialize};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a serializable struct.
    #[derive(Deserialize, Serialize)]
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

    // Default serialization with headers (fruit and cost).
    worksheet.deserialize_headers::<Produce>(0, 0)?;
    worksheet.serialize(&items)?;

    // Serialize the data but hide headers.
    let header_options = SerializeHeadersOptions::new().hide_headers(true);

    worksheet.deserialize_headers_with_options::<Produce>(0, 3, &header_options)?;
    worksheet.serialize(&items)?;

    // Save the file.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
