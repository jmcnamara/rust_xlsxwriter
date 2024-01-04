// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates different methods of handling custom
//! properties. The user can either merge them with the default properties or
//! use the custom properties exclusively.
//!
use rust_xlsxwriter::{CustomSerializeField, SerializeFieldOptions, Workbook, XlsxError};
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
        in_stock: bool,
    }

    // Create some data instances.
    let items = [
        Produce {
            fruit: "Peach",
            cost: 1.05,
            in_stock: true,
        },
        Produce {
            fruit: "Plum",
            cost: 0.15,
            in_stock: false,
        },
        Produce {
            fruit: "Pear",
            cost: 0.75,
            in_stock: true,
        },
    ];

    // Default handling of customized headers: the formatting is merged with the
    // default values so "in_stock" is still shown.
    let custom_headers = [
        CustomSerializeField::new("fruit").rename("Item"),
        CustomSerializeField::new("cost").rename("Price"),
        CustomSerializeField::new("in_stock").rename("Foo"),
    ];
    let header_options = SerializeFieldOptions::new().set_custom_headers(&custom_headers);

    worksheet.deserialize_headers_with_options::<Produce>(0, 0, &header_options)?;
    worksheet.serialize(&items)?;

    // Set the "use_custom_headers_only" option to shown only the specified
    // custom headers.
    let custom_headers = [
        CustomSerializeField::new("fruit").rename("Item"),
        CustomSerializeField::new("cost").rename("Price"),
    ];
    let header_options = SerializeFieldOptions::new()
        .set_custom_headers(&custom_headers)
        .use_custom_headers_only(true);

    worksheet.deserialize_headers_with_options::<Produce>(0, 4, &header_options)?;
    worksheet.serialize(&items)?;

    // This can also be used to set the order of the output.
    let custom_headers = [
        CustomSerializeField::new("cost").rename("Price"),
        CustomSerializeField::new("fruit").rename("Item"),
    ];
    let header_options = SerializeFieldOptions::new();
    let header_options = header_options
        .set_custom_headers(&custom_headers)
        .use_custom_headers_only(true);

    worksheet.deserialize_headers_with_options::<Produce>(0, 7, &header_options)?;
    worksheet.serialize(&items)?;

    // Save the file.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
