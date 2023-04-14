// custom - A module for creating the Excel Custom.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use chrono::{DateTime, Utc};

use crate::{xmlwriter::XMLWriter, CustomProperty, CustomPropertyType, DocProperties};

pub struct Custom {
    pub(crate) writer: XMLWriter,
    pub(crate) properties: DocProperties,
}

impl Custom {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    // Create a new Custom struct.
    pub(crate) fn new() -> Custom {
        let writer = XMLWriter::new();

        Custom {
            writer,

            properties: DocProperties::new(),
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the Properties element.
        self.write_properties();

        for (pid, property) in self.properties.custom_properties.clone().iter().enumerate() {
            // Write the property element.
            self.write_property(property, pid + 2);
        }

        // Close the Properties tag.
        self.writer.xml_end_tag("Properties");
    }

    // Write the <Properties> element.
    fn write_properties(&mut self) {
        let schema = "http://schemas.openxmlformats.org/officeDocument/2006".to_string();
        let xmlns = format!("{schema}/custom-properties");
        let xmlns_vt = format!("{schema}/docPropsVTypes");

        let attributes = [("xmlns", xmlns), ("xmlns:vt", xmlns_vt)];

        self.writer
            .xml_start_tag_with_attributes("Properties", &attributes);
    }

    // Write the <property> element.
    fn write_property(&mut self, property: &CustomProperty, pid: usize) {
        let fmtid = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}".to_string();
        let attributes = [
            ("fmtid", fmtid),
            ("pid", pid.to_string()),
            ("name", property.name.to_string()),
        ];

        self.writer
            .xml_start_tag_with_attributes("property", &attributes);

        match property.property_type {
            CustomPropertyType::Int => self.write_vt_i_4(property.number_int),
            CustomPropertyType::Bool => self.write_vt_bool(property.boolean),
            CustomPropertyType::Real => self.write_vt_r_8(property.number_real),
            CustomPropertyType::Text => self.write_vt_lpwstr(&property.text),
            CustomPropertyType::DateTime => self.write_vt_filetime(&property.datetime),
        }

        self.writer.xml_end_tag("property");
    }

    // Write the <vt:lpwstr> element.
    fn write_vt_lpwstr(&mut self, text: &str) {
        self.writer.xml_data_element("vt:lpwstr", text);
    }

    // Write the <vt:filetime> element.
    fn write_vt_filetime(&mut self, datetime: &DateTime<Utc>) {
        let utc_date = datetime.to_rfc3339_opts(chrono::SecondsFormat::Secs, true);

        self.writer.xml_data_element("vt:filetime", &utc_date);
    }

    // Write the <vt:i4> element.
    fn write_vt_i_4(&mut self, number: i32) {
        self.writer.xml_data_element("vt:i4", &number.to_string());
    }

    // Write the <vt:r8> element.
    fn write_vt_r_8(&mut self, number: f64) {
        self.writer.xml_data_element("vt:r8", &number.to_string());
    }

    // Write the <vt:bool> element.
    fn write_vt_bool(&mut self, boolean: bool) {
        self.writer
            .xml_data_element("vt:bool", &boolean.to_string());
    }
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use crate::custom::Custom;
    use crate::{test_functions::xml_to_vec, DocProperties};
    use chrono::{TimeZone, Utc};
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble1() {
        let mut custom = Custom::new();

        let properties = DocProperties::new().set_custom_property("Checked by", "Adam");

        custom.properties = properties;

        custom.assemble_xml_file();

        let got = custom.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
              <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Checked by">
                <vt:lpwstr>Adam</vt:lpwstr>
              </property>
            </Properties>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble2() {
        let date = Utc.with_ymd_and_hms(2016, 12, 12, 23, 0, 0).unwrap();

        let mut custom = Custom::new();

        let properties = DocProperties::new()
            .set_custom_property("Checked by", "Adam")
            .set_custom_property("Date completed", &date)
            .set_custom_property("Document number", 12345)
            .set_custom_property("Reference", 1.2345)
            .set_custom_property("Source", true)
            .set_custom_property("Status", false)
            .set_custom_property("Department", "Finance")
            .set_custom_property("Group", 1.2345678901234);

        custom.properties = properties;

        custom.assemble_xml_file();

        let got = custom.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
              <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Checked by">
                <vt:lpwstr>Adam</vt:lpwstr>
              </property>
              <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Date completed">
                <vt:filetime>2016-12-12T23:00:00Z</vt:filetime>
              </property>
              <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Document number">
                <vt:i4>12345</vt:i4>
              </property>
              <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="Reference">
                <vt:r8>1.2345</vt:r8>
              </property>
              <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="6" name="Source">
                <vt:bool>true</vt:bool>
              </property>
              <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="7" name="Status">
                <vt:bool>false</vt:bool>
              </property>
              <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="8" name="Department">
                <vt:lpwstr>Finance</vt:lpwstr>
              </property>
              <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="9" name="Group">
                <vt:r8>1.2345678901234</vt:r8>
              </property>
            </Properties>
            "#,
        );

        assert_eq!(expected, got);
    }
}
