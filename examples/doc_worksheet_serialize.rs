// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates serializing instances of a Serde derived
//! data structure to a worksheet.

use rust_xlsxwriter::{Format, Workbook, XlsxError};
use serde::Serialize;

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add a simple format for the headers.
    let format = Format::new().set_bold();

    // Create a serializable test struct.
    #[derive(Serialize)]
    #[serde(rename_all = "PascalCase")]
    struct Produce {
        fruit: &'static str,
        cost: f64,
    }

    // Create some data instances.
    let item1 = Produce {
        fruit: "Peach",
        cost: 1.05,
    };
    let item2 = Produce {
        fruit: "Plum",
        cost: 0.15,
    };
    let item3 = Produce {
        fruit: "Pear",
        cost: 0.75,
    };

    // Set up the start location and headers of the data to be serialized using
    // any temporary or valid instance.
    worksheet.serialize_headers_with_format(0, 0, &item1, &format)?;

    // Serialize the data.
    worksheet.serialize(&item1)?;
    worksheet.serialize(&item2)?;
    worksheet.serialize(&item3)?;

    // Save the file.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
