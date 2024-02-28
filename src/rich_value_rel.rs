// RichValueRel - A module for creating the Excel richValueRel.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

use crate::xmlwriter::XMLWriter;

pub struct RichValueRel {
    pub(crate) writer: XMLWriter,
    pub(crate) num_embedded_images: u32,
}

impl RichValueRel {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new RichValueRel struct.
    pub(crate) fn new() -> RichValueRel {
        let writer = XMLWriter::new();

        RichValueRel {
            writer,
            num_embedded_images: 0,
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the richValueRels element.
        self.write_rich_value_rels();

        // Close the final tag.
        self.writer.xml_end_tag("richValueRels");
    }

    // Write the <richValueRels> element.
    fn write_rich_value_rels(&mut self) {
        let attributes = [
            (
                "xmlns",
                "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel",
            ),
            (
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            ),
        ];

        self.writer.xml_start_tag("richValueRels", &attributes);

        for index in 1..=self.num_embedded_images {
            // Write the rel element.
            self.write_rel(index);
        }
    }

    // Write the <rel> element.
    fn write_rel(&mut self, index: u32) {
        let attributes = [("r:id", format!("rId{index}"))];

        self.writer.xml_empty_tag("rel", &attributes);
    }
}
