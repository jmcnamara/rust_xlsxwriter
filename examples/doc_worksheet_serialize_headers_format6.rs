// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates serializing instances of a Serde derived
//! data structure to a worksheet with header and value formatting.

use rust_xlsxwriter::{
    CustomSerializeHeader, Format, FormatBorder, SerializeHeadersOptions, Workbook, XlsxError,
};
use serde::{Deserialize, Serialize};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add some formats to use with the serialization data.
    let header_format = Format::new()
        .set_bold()
        .set_border(FormatBorder::Thin)
        .set_background_color("C6EFCE");

    let currency_format = Format::new().set_num_format("$0.00");

    // Create a serializable struct.
    #[derive(Deserialize, Serialize)]
    struct Produce {
        #[serde(rename = "Item")]
        fruit: &'static str,

        #[serde(rename = "Price")]
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
    let custom_headers = [CustomSerializeHeader::new("Price").set_value_format(&currency_format)];

    let header_options = SerializeHeadersOptions::new()
        .set_header_format(&header_format)
        .set_custom_headers(&custom_headers);

    // Set the serialization location and custom headers.
    worksheet.deserialize_headers_with_options::<Produce>(1, 1, &header_options)?;

    // Serialize the data.
    worksheet.serialize(&item1)?;
    worksheet.serialize(&item2)?;
    worksheet.serialize(&item3)?;

    // Save the file.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
