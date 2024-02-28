// RichValueTypes - A module for creating the Excel rdRichValueTypes.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

use crate::xmlwriter::XMLWriter;

pub struct RichValueTypes {
    pub(crate) writer: XMLWriter,
}

impl RichValueTypes {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new RichValueTypes struct.
    pub(crate) fn new() -> RichValueTypes {
        let writer = XMLWriter::new();

        RichValueTypes { writer }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the rvTypesInfo element.
        self.write_rv_types_info();

        // Write the global element.
        self.write_global();

        // Close the final tag.
        self.writer.xml_end_tag("rvTypesInfo");
    }

    // Write the <rvTypesInfo> element.
    fn write_rv_types_info(&mut self) {
        let attributes = [
            (
                "xmlns",
                "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2",
            ),
            (
                "xmlns:mc",
                "http://schemas.openxmlformats.org/markup-compatibility/2006",
            ),
            ("mc:Ignorable", "x"),
            (
                "xmlns:x",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            ),
        ];

        self.writer.xml_start_tag("rvTypesInfo", &attributes);
    }

    // Write the <global> element.
    fn write_global(&mut self) {
        let key_flags = [
            ("_Self", "ExcludeFromFile", "ExcludeFromCalcComparison"),
            ("_DisplayString", "ExcludeFromCalcComparison", ""),
            ("_Flags", "ExcludeFromCalcComparison", ""),
            ("_Format", "ExcludeFromCalcComparison", ""),
            ("_SubLabel", "ExcludeFromCalcComparison", ""),
            ("_Attribution", "ExcludeFromCalcComparison", ""),
            ("_Icon", "ExcludeFromCalcComparison", ""),
            ("_Display", "ExcludeFromCalcComparison", ""),
            ("_CanonicalPropertyNames", "ExcludeFromCalcComparison", ""),
            ("_ClassificationId", "ExcludeFromCalcComparison", ""),
        ];

        self.writer.xml_start_tag_only("global");
        self.writer.xml_start_tag_only("keyFlags");

        for (key, flag1, flag2) in key_flags {
            self.write_key(key);
            self.write_flag(flag1);

            if !flag2.is_empty() {
                self.write_flag(flag2);
            }

            self.writer.xml_end_tag("key");
        }

        self.writer.xml_end_tag("keyFlags");
        self.writer.xml_end_tag("global");
    }

    // Write the <key> element.
    fn write_key(&mut self, name: &str) {
        let attributes = [("name", name)];

        self.writer.xml_start_tag("key", &attributes);
    }

    // Write the <flag> element.
    fn write_flag(&mut self, name: &str) {
        let attributes = [("name", name), ("value", "1")];

        self.writer.xml_empty_tag("flag", &attributes);
    }
}
