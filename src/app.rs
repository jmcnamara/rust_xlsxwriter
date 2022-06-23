// app - A module for creating the Excel app.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use crate::xmlwriter::XMLWriter;

pub struct App<'a> {
    pub writer: &'a mut XMLWriter<'a>,
    heading_pairs: Vec<(String, u16)>,
    table_parts: Vec<String>,
    doc_security: u8,
}

impl<'a> App<'a> {
    // Create a new App struct.
    pub fn new(writer: &'a mut XMLWriter<'a>) -> App<'a> {
        App {
            writer,
            heading_pairs: vec![],
            table_parts: vec![],
            doc_security: 0,
        }
    }

    pub fn add_heading_pair(&mut self, key: &str, value: u16) {
        self.heading_pairs.push((key.to_string(), value));
    }

    pub fn add_part_name(&mut self, part_name: &str) {
        self.table_parts.push(part_name.to_string());
    }

    //  Assemble and write the XML file.
    pub fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the Properties element.
        self.write_properties();

        // Write the Application element.
        self.write_application();

        // Write the DocSecurity element.
        self.write_doc_security();

        // Write the ScaleCrop element.
        self.write_scale_crop();

        // Write the HeadingPairs element.
        self.write_heading_pairs();

        // Write the TitlesOfParts element.
        self.write_titles_of_parts();

        // Write the Company element.
        self.write_company();

        // Write the LinksUpToDate element.
        self.write_links_up_to_date();

        // Write the SharedDoc element.
        self.write_shared_doc();

        // Write the HyperlinksChanged element.
        self.write_hyperlinks_changed();

        // Write the AppVersion element.
        self.write_app_version();

        // Close the Properties tag.
        self.writer.xml_end_tag("Properties");
    }

    // Write the <Properties> element.
    fn write_properties(&mut self) {
        let xmlns = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        let xmlns_vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

        let attributes = vec![("xmlns", xmlns), ("xmlns:vt", xmlns_vt)];

        self.writer.xml_start_tag_attr("Properties", &attributes);
    }

    // Write the <Application> element.
    fn write_application(&mut self) {
        self.writer
            .xml_data_element("Application", "Microsoft Excel");
    }

    // Write the <DocSecurity> element.
    fn write_doc_security(&mut self) {
        self.writer
            .xml_data_element("DocSecurity", &self.doc_security.to_string());
    }

    // Write the <ScaleCrop> element.
    fn write_scale_crop(&mut self) {
        self.writer.xml_data_element("ScaleCrop", "false");
    }

    // Write the <HeadingPairs> element.
    fn write_heading_pairs(&mut self) {
        self.writer.xml_start_tag("HeadingPairs");

        // Write the vt:vector element for headings.
        self.write_heading_vector();

        self.writer.xml_end_tag("HeadingPairs");
    }

    // Write the <vt:vector> element.
    fn write_heading_vector(&mut self) {
        let size = self.heading_pairs.len() * 2;
        let size = size.to_string();
        let attributes = vec![("size", size.as_str()), ("baseType", "variant")];

        self.writer.xml_start_tag_attr("vt:vector", &attributes);

        for heading_pair in self.heading_pairs.clone() {
            self.writer.xml_start_tag("vt:variant");
            self.write_vt_lpstr(&heading_pair.0);
            self.writer.xml_end_tag("vt:variant");

            self.writer.xml_start_tag("vt:variant");
            self.write_vt_i4(heading_pair.1);
            self.writer.xml_end_tag("vt:variant");
        }

        self.writer.xml_end_tag("vt:vector");
    }

    // Write the <TitlesOfParts> element.
    fn write_titles_of_parts(&mut self) {
        self.writer.xml_start_tag("TitlesOfParts");

        self.write_title_parts_vector();

        self.writer.xml_end_tag("TitlesOfParts");
    }

    // Write the <vt:vector> element.
    fn write_title_parts_vector(&mut self) {
        let size = self.table_parts.len();
        let size = size.to_string();
        let attributes = vec![("size", size.as_str()), ("baseType", "lpstr")];

        self.writer.xml_start_tag_attr("vt:vector", &attributes);

        for part_name in self.table_parts.clone() {
            self.write_vt_lpstr(&part_name);
        }

        self.writer.xml_end_tag("vt:vector");
    }

    // Write the <vt:lpstr> element.
    fn write_vt_lpstr(&mut self, data: &str) {
        self.writer.xml_data_element("vt:lpstr", data);
    }

    // Write the <vt:i4> element.
    fn write_vt_i4(&mut self, count: u16) {
        self.writer.xml_data_element("vt:i4", &count.to_string());
    }

    // Write the <Company> element.
    fn write_company(&mut self) {
        self.writer.xml_data_element("Company", "");
    }

    // Write the <LinksUpToDate> element.
    fn write_links_up_to_date(&mut self) {
        self.writer.xml_data_element("LinksUpToDate", "false");
    }

    // Write the <SharedDoc> element.
    fn write_shared_doc(&mut self) {
        self.writer.xml_data_element("SharedDoc", "false");
    }

    // Write the <HyperlinksChanged> element.
    fn write_hyperlinks_changed(&mut self) {
        self.writer.xml_data_element("HyperlinksChanged", "false");
    }

    // Write the <AppVersion> element.
    fn write_app_version(&mut self) {
        self.writer.xml_data_element("AppVersion", "12.0000");
    }
}

#[cfg(test)]
mod tests {

    use super::App;
    use super::XMLWriter;
    use crate::test_functions::read_xmlfile_data;
    use crate::test_functions::xml_to_vec;

    use pretty_assertions::assert_eq;
    use tempfile::tempfile;

    #[test]
    fn test_assemble1() {
        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        let mut app = App::new(&mut writer);

        app.add_heading_pair("Worksheets", 1);
        app.add_part_name("Sheet1");

        app.assemble_xml_file();

        let got = read_xmlfile_data(&mut tempfile);
        let got = xml_to_vec(&got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
                    <Application>Microsoft Excel</Application>
                    <DocSecurity>0</DocSecurity>
                    <ScaleCrop>false</ScaleCrop>
                    <HeadingPairs>
                        <vt:vector size="2" baseType="variant">
                        <vt:variant>
                            <vt:lpstr>Worksheets</vt:lpstr>
                        </vt:variant>
                        <vt:variant>
                            <vt:i4>1</vt:i4>
                        </vt:variant>
                        </vt:vector>
                    </HeadingPairs>
                    <TitlesOfParts>
                        <vt:vector size="1" baseType="lpstr">
                        <vt:lpstr>Sheet1</vt:lpstr>
                        </vt:vector>
                    </TitlesOfParts>
                    <Company>
                    </Company>
                    <LinksUpToDate>false</LinksUpToDate>
                    <SharedDoc>false</SharedDoc>
                    <HyperlinksChanged>false</HyperlinksChanged>
                    <AppVersion>12.0000</AppVersion>
                </Properties>
                "#,
        );

        assert_eq!(got, expected);
    }

    #[test]
    fn test_assemble2() {
        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        let mut app = App::new(&mut writer);

        app.add_heading_pair("Worksheets", 2);
        app.add_part_name("Sheet1");
        app.add_part_name("Sheet2");

        app.assemble_xml_file();

        let got = read_xmlfile_data(&mut tempfile);
        let got = xml_to_vec(&got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
                <Application>Microsoft Excel</Application>
                <DocSecurity>0</DocSecurity>
                <ScaleCrop>false</ScaleCrop>
                <HeadingPairs>
                    <vt:vector size="2" baseType="variant">
                    <vt:variant>
                        <vt:lpstr>Worksheets</vt:lpstr>
                    </vt:variant>
                    <vt:variant>
                        <vt:i4>2</vt:i4>
                    </vt:variant>
                    </vt:vector>
                </HeadingPairs>
                <TitlesOfParts>
                    <vt:vector size="2" baseType="lpstr">
                    <vt:lpstr>Sheet1</vt:lpstr>
                    <vt:lpstr>Sheet2</vt:lpstr>
                    </vt:vector>
                </TitlesOfParts>
                <Company>
                </Company>
                <LinksUpToDate>false</LinksUpToDate>
                <SharedDoc>false</SharedDoc>
                <HyperlinksChanged>false</HyperlinksChanged>
                <AppVersion>12.0000</AppVersion>
                </Properties>
                "#,
        );

        assert_eq!(got, expected);
    }

    #[test]
    fn test_assemble3() {
        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        let mut app = App::new(&mut writer);

        app.add_heading_pair("Worksheets", 1);
        app.add_heading_pair("Named Ranges", 1);
        app.add_part_name("Sheet1");
        app.add_part_name("Sheet1!Print_Titles");

        app.assemble_xml_file();

        let got = read_xmlfile_data(&mut tempfile);
        let got = xml_to_vec(&got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
                <Application>Microsoft Excel</Application>
                <DocSecurity>0</DocSecurity>
                <ScaleCrop>false</ScaleCrop>
                <HeadingPairs>
                    <vt:vector size="4" baseType="variant">
                    <vt:variant>
                        <vt:lpstr>Worksheets</vt:lpstr>
                    </vt:variant>
                    <vt:variant>
                        <vt:i4>1</vt:i4>
                    </vt:variant>
                    <vt:variant>
                        <vt:lpstr>Named Ranges</vt:lpstr>
                    </vt:variant>
                    <vt:variant>
                        <vt:i4>1</vt:i4>
                    </vt:variant>
                    </vt:vector>
                </HeadingPairs>
                <TitlesOfParts>
                    <vt:vector size="2" baseType="lpstr">
                    <vt:lpstr>Sheet1</vt:lpstr>
                    <vt:lpstr>Sheet1!Print_Titles</vt:lpstr>
                    </vt:vector>
                </TitlesOfParts>
                <Company>
                </Company>
                <LinksUpToDate>false</LinksUpToDate>
                <SharedDoc>false</SharedDoc>
                <HyperlinksChanged>false</HyperlinksChanged>
                <AppVersion>12.0000</AppVersion>
                </Properties>
                "#,
        );

        assert_eq!(got, expected);
    }
}
