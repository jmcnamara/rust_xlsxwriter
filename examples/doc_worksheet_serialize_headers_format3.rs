// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates formatting cells during serialization.
//!
use rust_xlsxwriter::{
    CustomSerializeHeader, Format, SerializeHeadersOptions, Workbook, XlsxError,
};
use serde::Serialize;

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set a header format.
    let cell_format = Format::new().set_num_format("$0.00");

    // Create a serializable test struct.
    #[derive(Serialize)]
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

    // Set the serialization location and headers.
    let custom_headers = [CustomSerializeHeader::new("cost").set_cell_format(&cell_format)];
    let header_options = SerializeHeadersOptions::new().set_custom_headers(&custom_headers);

    // Set the serialization location and custom headers.
    worksheet.serialize_headers_with_options(0, 0, &item1, &header_options)?;

    // Serialize the data.
    worksheet.serialize(&item1)?;
    worksheet.serialize(&item2)?;
    worksheet.serialize(&item3)?;

    // Save the file.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
