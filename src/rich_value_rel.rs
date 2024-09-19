// RichValueRel - A module for creating the Excel richValueRel.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

use std::io::Cursor;

use crate::xmlwriter::{xml_declaration, xml_empty_tag, xml_end_tag, xml_start_tag};

pub struct RichValueRel {
    pub(crate) writer: Cursor<Vec<u8>>,
    pub(crate) num_embedded_images: u32,
}

impl RichValueRel {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new RichValueRel struct.
    pub(crate) fn new() -> RichValueRel {
        let writer = Cursor::new(Vec::with_capacity(2048));

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
        xml_declaration(&mut self.writer);

        // Write the richValueRels element.
        self.write_rich_value_rels();

        // Close the final tag.
        xml_end_tag(&mut self.writer, "richValueRels");
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

        xml_start_tag(&mut self.writer, "richValueRels", &attributes);

        for index in 1..=self.num_embedded_images {
            // Write the rel element.
            self.write_rel(index);
        }
    }

    // Write the <rel> element.
    fn write_rel(&mut self, index: u32) {
        let attributes = [("r:id", format!("rId{index}"))];

        xml_empty_tag(&mut self.writer, "rel", &attributes);
    }
}
