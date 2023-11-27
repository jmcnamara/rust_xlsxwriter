// Custom unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod custom_tests {

    use crate::custom::Custom;
    use crate::ExcelDateTime;
    use crate::{test_functions::xml_to_vec, DocProperties};
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
        let date = ExcelDateTime::from_ymd(2016, 12, 12)
            .unwrap()
            .and_hms(23, 0, 0)
            .unwrap();

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
