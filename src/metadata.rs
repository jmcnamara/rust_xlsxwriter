// metadata - A module for creating the Excel Metadata.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

use crate::xmlwriter::XMLWriter;

pub struct Metadata {
    pub(crate) writer: XMLWriter,
}

impl Metadata {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new Metadata struct.
    pub fn new() -> Metadata {
        let writer = XMLWriter::new();

        Metadata { writer }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the metadata element.
        self.write_metadata();

        // Write the metadataTypes element.
        self.write_metadata_types();

        // Write the futureMetadata element.
        self.write_future_metadata();

        // Write the cellMetadata element.
        self.write_cell_metadata();

        // Close the metadata tag.
        self.writer.xml_end_tag("metadata");
    }

    // Write the <metadata> element.
    fn write_metadata(&mut self) {
        let attributes = [
            (
                "xmlns",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            ),
            (
                "xmlns:xda",
                "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray",
            ),
        ];

        self.writer.xml_start_tag("metadata", &attributes);
    }

    // Write the <metadataTypes> element.
    fn write_metadata_types(&mut self) {
        let attributes = [("count", "1")];

        self.writer.xml_start_tag("metadataTypes", &attributes);

        // Write the metadataType element.
        self.write_metadata_type();

        self.writer.xml_end_tag("metadataTypes");
    }

    // Write the <metadataType> element.
    fn write_metadata_type(&mut self) {
        let attributes = [
            ("name", "XLDAPR"),
            ("minSupportedVersion", "120000"),
            ("copy", "1"),
            ("pasteAll", "1"),
            ("pasteValues", "1"),
            ("merge", "1"),
            ("splitFirst", "1"),
            ("rowColShift", "1"),
            ("clearFormats", "1"),
            ("clearComments", "1"),
            ("assign", "1"),
            ("coerce", "1"),
            ("cellMeta", "1"),
        ];

        self.writer.xml_empty_tag("metadataType", &attributes);
    }

    // Write the <futureMetadata> element.
    fn write_future_metadata(&mut self) {
        let attributes = [("name", "XLDAPR"), ("count", "1")];

        self.writer.xml_start_tag("futureMetadata", &attributes);
        self.writer.xml_start_tag_only("bk");
        self.writer.xml_start_tag_only("extLst");

        // Write the ext element.
        self.write_ext();

        self.writer.xml_end_tag("extLst");
        self.writer.xml_end_tag("bk");
        self.writer.xml_end_tag("futureMetadata");
    }

    // Write the <ext> element.
    fn write_ext(&mut self) {
        let attributes = [("uri", "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}")];

        self.writer.xml_start_tag("ext", &attributes);

        // Write the xda:dynamicArrayProperties element.
        self.write_xda_dynamic_array_properties();

        self.writer.xml_end_tag("ext");
    }

    // Write the <xda:dynamicArrayProperties> element.
    fn write_xda_dynamic_array_properties(&mut self) {
        let attributes = [("fDynamic", "1"), ("fCollapsed", "0")];

        self.writer
            .xml_empty_tag("xda:dynamicArrayProperties", &attributes);
    }

    // Write the <cellMetadata> element.
    fn write_cell_metadata(&mut self) {
        let attributes = [("count", "1")];

        self.writer.xml_start_tag("cellMetadata", &attributes);
        self.writer.xml_start_tag_only("bk");

        // Write the rc element.
        self.write_rc();

        self.writer.xml_end_tag("bk");
        self.writer.xml_end_tag("cellMetadata");
    }

    // Write the <rc> element.
    fn write_rc(&mut self) {
        let attributes = [("t", "1"), ("v", "0")];

        self.writer.xml_empty_tag("rc", &attributes);
    }
}
