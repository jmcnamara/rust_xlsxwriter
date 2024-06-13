// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example adding a Sensitivity Label to an Excel file using custom document
//! properties. See the main docs for an explanation of how to extract the
//! metadata.

use rust_xlsxwriter::{DocProperties, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Metadata extracted from a company specific file.
    let site_id = "cb46c030-1825-4e81-a295-151c039dbf02";
    let action_id = "88124cf5-1340-457d-90e1-0000a9427c99";
    let company_guid = "2096f6a2-d2f7-48be-b329-b73aaa526e5d";

    // Add the document properties. Note that these should all be in text format.
    let properties = DocProperties::new()
        .set_custom_property(format!("MSIP_Label_{company_guid}_Method"), "Privileged")
        .set_custom_property(format!("MSIP_Label_{company_guid}_Name"), "Confidential")
        .set_custom_property(format!("MSIP_Label_{company_guid}_SiteId"), site_id)
        .set_custom_property(format!("MSIP_Label_{company_guid}_ActionId"), action_id)
        .set_custom_property(format!("MSIP_Label_{company_guid}_ContentBits"), "2");

    workbook.set_properties(&properties);

    workbook.save("sensitivity_label.xlsx")?;

    Ok(())
}
