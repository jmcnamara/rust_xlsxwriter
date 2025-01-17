// RichValueTypes - A module for creating the Excel rdRichValueTypes.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

use std::io::Cursor;

use crate::xmlwriter::{
    xml_declaration, xml_empty_tag, xml_end_tag, xml_start_tag, xml_start_tag_only,
};

pub struct RichValueTypes {
    pub(crate) writer: Cursor<Vec<u8>>,
}

impl RichValueTypes {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new RichValueTypes struct.
    pub(crate) fn new() -> RichValueTypes {
        let writer = Cursor::new(Vec::with_capacity(2048));

        RichValueTypes { writer }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        // Write the rvTypesInfo element.
        self.write_rv_types_info();

        // Write the global element.
        self.write_global();

        // Close the final tag.
        xml_end_tag(&mut self.writer, "rvTypesInfo");
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

        xml_start_tag(&mut self.writer, "rvTypesInfo", &attributes);
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

        xml_start_tag_only(&mut self.writer, "global");
        xml_start_tag_only(&mut self.writer, "keyFlags");

        for (key, flag1, flag2) in key_flags {
            self.write_key(key);
            self.write_flag(flag1);

            if !flag2.is_empty() {
                self.write_flag(flag2);
            }

            xml_end_tag(&mut self.writer, "key");
        }

        xml_end_tag(&mut self.writer, "keyFlags");
        xml_end_tag(&mut self.writer, "global");
    }

    // Write the <key> element.
    fn write_key(&mut self, name: &str) {
        let attributes = [("name", name)];

        xml_start_tag(&mut self.writer, "key", &attributes);
    }

    // Write the <flag> element.
    fn write_flag(&mut self, name: &str) {
        let attributes = [("name", name), ("value", "1")];

        xml_empty_tag(&mut self.writer, "flag", &attributes);
    }
}
