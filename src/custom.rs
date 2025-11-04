// custom - A module for creating the Excel Custom.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

mod tests;

use std::io::Cursor;

use crate::xmlwriter::{xml_data_element_only, xml_declaration, xml_end_tag, xml_start_tag};
use crate::{CustomProperty, CustomPropertyType, DocProperties};

pub struct Custom {
    pub(crate) writer: Cursor<Vec<u8>>,
    pub(crate) properties: DocProperties,
}

impl Custom {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    // Create a new Custom struct.
    pub(crate) fn new() -> Custom {
        let writer = Cursor::new(Vec::with_capacity(2048));

        Custom {
            writer,

            properties: DocProperties::new(),
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and generate the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        // Write the <Properties> element.
        self.write_properties();

        for (pid, property) in self.properties.custom_properties.clone().iter().enumerate() {
            // Write the <property> element.
            self.write_property(property, pid + 2);
        }

        // Close the final tag.
        xml_end_tag(&mut self.writer, "Properties");
    }

    // Write the <Properties> element.
    fn write_properties(&mut self) {
        let schema = "http://schemas.openxmlformats.org/officeDocument/2006".to_string();
        let xmlns = format!("{schema}/custom-properties");
        let xmlns_vt = format!("{schema}/docPropsVTypes");

        let attributes = [("xmlns", xmlns), ("xmlns:vt", xmlns_vt)];

        xml_start_tag(&mut self.writer, "Properties", &attributes);
    }

    // Write the <property> element.
    fn write_property(&mut self, property: &CustomProperty, pid: usize) {
        let fmtid = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}".to_string();
        let attributes = [
            ("fmtid", fmtid),
            ("pid", pid.to_string()),
            ("name", property.name.clone()),
        ];

        xml_start_tag(&mut self.writer, "property", &attributes);

        match property.property_type {
            CustomPropertyType::Int => self.write_vt_i_4(property.number_int),
            CustomPropertyType::Bool => self.write_vt_bool(property.boolean),
            CustomPropertyType::Real => self.write_vt_r_8(property.number_real),
            CustomPropertyType::Text => self.write_vt_lpwstr(&property.text),
            CustomPropertyType::DateTime => self.write_vt_filetime(&property.datetime),
        }

        xml_end_tag(&mut self.writer, "property");
    }

    // Write the <vt:lpwstr> element.
    fn write_vt_lpwstr(&mut self, text: &str) {
        xml_data_element_only(&mut self.writer, "vt:lpwstr", text);
    }

    // Write the <vt:filetime> element.
    fn write_vt_filetime(&mut self, utc_datetime: &str) {
        xml_data_element_only(&mut self.writer, "vt:filetime", utc_datetime);
    }

    // Write the <vt:i4> element.
    fn write_vt_i_4(&mut self, number: i32) {
        xml_data_element_only(&mut self.writer, "vt:i4", &number.to_string());
    }

    // Write the <vt:r8> element.
    fn write_vt_r_8(&mut self, number: f64) {
        xml_data_element_only(&mut self.writer, "vt:r8", &number.to_string());
    }

    // Write the <vt:bool> element.
    fn write_vt_bool(&mut self, boolean: bool) {
        xml_data_element_only(&mut self.writer, "vt:bool", &boolean.to_string());
    }
}
