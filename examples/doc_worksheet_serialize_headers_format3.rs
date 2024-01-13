// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates formatting cells during serialization.
//!
use rust_xlsxwriter::{CustomSerializeField, Format, SerializeFieldOptions, Workbook, XlsxError};
use serde::{Deserialize, Serialize};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set a cell value format.
    let value_format = Format::new().set_num_format("$0.00");

    // Create a serializable struct.
    #[derive(Deserialize, Serialize)]
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

    // Set the custom headers.
    let header_options = SerializeFieldOptions::new()
        .set_custom_headers(&[CustomSerializeField::new("cost").set_value_format(value_format)]);

    // Set the serialization location and headers.
    worksheet.deserialize_headers_with_options::<Produce>(0, 0, &header_options)?;

    // Serialize the data.
    worksheet.serialize(&item1)?;
    worksheet.serialize(&item2)?;
    worksheet.serialize(&item3)?;

    // Save the file.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
