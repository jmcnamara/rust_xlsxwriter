// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates serializing instances of a Serde derived
//! data structure to a worksheet with a user defined worksheet table.

use rust_xlsxwriter::{
    SerializeFieldOptions, Table, TableColumn, TableFunction, Workbook, XlsxError,
};
use serde::{Deserialize, Serialize};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a serializable struct.
    #[derive(Deserialize, Serialize)]
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

    // Set the caption and subtotal in the total row.
    let columns = vec![
        TableColumn::new().set_total_label("Total"),
        TableColumn::new().set_total_function(TableFunction::Sum),
    ];

    // Create a new table and configure the total row.
    let table = Table::new().set_total_row(true).set_columns(&columns);

    // Set the header options.
    let header_options = SerializeFieldOptions::new().set_table(table);

    // Set the serialization location and custom headers.
    worksheet.deserialize_headers_with_options::<Produce>(0, 0, &header_options)?;

    // Serialize the data.
    worksheet.serialize(&item1)?;
    worksheet.serialize(&item2)?;
    worksheet.serialize(&item3)?;

    // Save the file.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
