// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates serializing instances of a Serde derived
//! data structure to a worksheet using different methods (both serialization
//! and deserialization).

use rust_xlsxwriter::{
    CustomSerializeHeader, Format, FormatBorder, SerializeHeadersOptions, Workbook, XlsxError,
};
use serde::{Deserialize, Serialize};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set some column widths for clarity.
    worksheet.set_column_width(2, 4)?;
    worksheet.set_column_width(5, 4)?;
    worksheet.set_column_width(8, 4)?;

    // Add some formats to use with the serialization data.
    let header_format = Format::new()
        .set_bold()
        .set_border(FormatBorder::Thin)
        .set_background_color("C6EFCE");

    let currency_format = Format::new().set_num_format("$0.00");

    // Create a serializable test struct.
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

    // 1. Set the serialization location and headers with `deserialize_headers()`
    //    and serialize some data.
    worksheet.deserialize_headers::<Produce>(0, 0)?;

    worksheet.serialize(&item1)?;
    worksheet.serialize(&item2)?;
    worksheet.serialize(&item3)?;

    // 2. Set the serialization location and formatted headers with
    //    `serialize_headers_with_format()` and serialize some data.
    worksheet.serialize_headers_with_format(0, 3, &item1, &header_format)?;

    worksheet.serialize(&item1)?;
    worksheet.serialize(&item2)?;
    worksheet.serialize(&item3)?;

    // 3. Set the serialization location and headers with custom headers. We use
    //    the customization to set the header format and also the cell format
    //    for the number values.
    let custom_headers = [
        CustomSerializeHeader::new("fruit").rename("Item"),
        CustomSerializeHeader::new("cost")
            .rename("Price")
            .set_cell_format(&currency_format),
    ];
    let header_options = SerializeHeadersOptions::new()
        .set_header_format(&header_format)
        .set_custom_headers(&custom_headers);

    // Set the serialization location and custom headers.
    worksheet.deserialize_headers_with_options::<Produce>(0, 6, &header_options)?;

    worksheet.serialize(&item1)?;
    worksheet.serialize(&item2)?;
    worksheet.serialize(&item3)?;

    // 4. Set the serialization location and headers with custom options. We use
    //    the customization to turn off the headers.
    let header_options = SerializeHeadersOptions::new().hide_headers(true);

    // Set the serialization location and custom headers.
    worksheet.serialize_headers_with_options(0, 9, &item1, &header_options)?;

    worksheet.serialize(&item1)?;
    worksheet.serialize(&item2)?;
    worksheet.serialize(&item3)?;

    // Save the file.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
