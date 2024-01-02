// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Create a simple workbook to demonstrate a constant checksum due to the a
//! constant creation date.

use chrono::{TimeZone, Utc};
use rust_xlsxwriter::{DocProperties, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Create a file creation date for the file.
    let date = Utc.with_ymd_and_hms(2023, 1, 1, 0, 0, 0).unwrap();

    // Add it to the document metadata.
    let properties = DocProperties::new().set_creation_datetime(&date);
    workbook.set_properties(&properties);

    let worksheet = workbook.add_worksheet();
    worksheet.write_string(0, 0, "Hello")?;

    workbook.save("properties.xlsx")?;

    Ok(())
}
