// RichValueStructure - A module for creating the Excel rdrichvaluestructure.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

use crate::xmlwriter::XMLWriter;

pub struct RichValueStructure {
    pub(crate) writer: XMLWriter,
    pub(crate) has_embedded_image_descriptions: bool,
}

impl RichValueStructure {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new RichValueStructure struct.
    pub(crate) fn new() -> RichValueStructure {
        let writer = XMLWriter::new();

        RichValueStructure {
            writer,
            has_embedded_image_descriptions: false,
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the rvStructures element.
        self.write_rv_structures();

        // Close the final tag.
        self.writer.xml_end_tag("rvStructures");
    }

    // Write the <rvStructures> element.
    fn write_rv_structures(&mut self) {
        let attributes = [
            (
                "xmlns",
                "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata",
            ),
            ("count", "1"),
        ];

        self.writer.xml_start_tag("rvStructures", &attributes);

        // Write the s element.
        self.write_s();
    }

    // Write the <s> element.
    fn write_s(&mut self) {
        let attributes = [("t", "_localImage")];

        self.writer.xml_start_tag("s", &attributes);

        // Write the k elements.
        self.write_k("_rvRel:LocalImageIdentifier", "i");
        self.write_k("CalcOrigin", "i");

        if self.has_embedded_image_descriptions {
            self.write_k("Text", "s");
        }

        self.writer.xml_end_tag("s");
    }

    // Write the <k> element.
    fn write_k(&mut self, name: &str, name_type: &str) {
        let attributes = [("n", name), ("t", name_type)];

        self.writer.xml_empty_tag("k", &attributes);
    }
}
