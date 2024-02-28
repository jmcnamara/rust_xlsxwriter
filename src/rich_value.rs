// RichValue - A module for creating the Excel rdrichvalue.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

use crate::{xmlwriter::XMLWriter, Image};

pub struct RichValue<'a> {
    pub(crate) writer: XMLWriter,
    pub(crate) embedded_images: &'a Vec<Image>,
}

impl<'a> RichValue<'a> {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new RichValue struct.
    pub(crate) fn new(embedded_images: &Vec<Image>) -> RichValue {
        let writer = XMLWriter::new();

        RichValue {
            writer,
            embedded_images,
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the rvData element.
        self.write_rv_data();

        // Close the final tag.
        self.writer.xml_end_tag("rvData");
    }

    // Write the <rvData> element.
    fn write_rv_data(&mut self) {
        let attributes = [
            (
                "xmlns",
                "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata".to_string(),
            ),
            ("count", self.embedded_images.len().to_string()),
        ];

        self.writer.xml_start_tag("rvData", &attributes);

        for (index, image) in self.embedded_images.iter().enumerate() {
            // Write the rv element.
            self.write_rv(index, image);
        }
    }

    // Write the <rv> element.
    fn write_rv(&mut self, index: usize, image: &Image) {
        let attributes = [("s", "0")];
        let mut value = "5";

        if image.decorative {
            value = "6";
        }

        self.writer.xml_start_tag("rv", &attributes);

        // Write the v element.
        self.write_v(&index.to_string());
        self.write_v(value);

        if !image.alt_text.is_empty() {
            self.write_v(&image.alt_text);
        }

        self.writer.xml_end_tag("rv");
    }

    // Write the <v> element.
    fn write_v(&mut self, value: &str) {
        self.writer.xml_data_element_only("v", value);
    }
}
