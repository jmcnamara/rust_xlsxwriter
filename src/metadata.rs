// metadata - A module for creating the Excel Metadata.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

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

    //  Assemble and write the XML file.
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
        let attributes = vec![
            (
                "xmlns",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
            ),
            (
                "xmlns:xda",
                "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray".to_string(),
            ),
        ];

        self.writer.xml_start_tag_attr("metadata", &attributes);
    }

    // Write the <metadataTypes> element.
    fn write_metadata_types(&mut self) {
        let attributes = vec![("count", "1".to_string())];

        self.writer.xml_start_tag_attr("metadataTypes", &attributes);

        // Write the metadataType element.
        self.write_metadata_type();

        self.writer.xml_end_tag("metadataTypes");
    }

    // Write the <metadataType> element.
    fn write_metadata_type(&mut self) {
        let attributes = vec![
            ("name", "XLDAPR".to_string()),
            ("minSupportedVersion", "120000".to_string()),
            ("copy", "1".to_string()),
            ("pasteAll", "1".to_string()),
            ("pasteValues", "1".to_string()),
            ("merge", "1".to_string()),
            ("splitFirst", "1".to_string()),
            ("rowColShift", "1".to_string()),
            ("clearFormats", "1".to_string()),
            ("clearComments", "1".to_string()),
            ("assign", "1".to_string()),
            ("coerce", "1".to_string()),
            ("cellMeta", "1".to_string()),
        ];

        self.writer.xml_empty_tag_attr("metadataType", &attributes);
    }

    // Write the <futureMetadata> element.
    fn write_future_metadata(&mut self) {
        let attributes = vec![("name", "XLDAPR".to_string()), ("count", "1".to_string())];

        self.writer
            .xml_start_tag_attr("futureMetadata", &attributes);
        self.writer.xml_start_tag("bk");
        self.writer.xml_start_tag("extLst");

        // Write the ext element.
        self.write_ext();

        self.writer.xml_end_tag("extLst");
        self.writer.xml_end_tag("bk");
        self.writer.xml_end_tag("futureMetadata");
    }

    // Write the <ext> element.
    fn write_ext(&mut self) {
        let attributes = vec![("uri", "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}".to_string())];

        self.writer.xml_start_tag_attr("ext", &attributes);

        // Write the xda:dynamicArrayProperties element.
        self.write_xda_dynamic_array_properties();

        self.writer.xml_end_tag("ext");
    }

    // Write the <xda:dynamicArrayProperties> element.
    fn write_xda_dynamic_array_properties(&mut self) {
        let attributes = vec![
            ("fDynamic", "1".to_string()),
            ("fCollapsed", "0".to_string()),
        ];

        self.writer
            .xml_empty_tag_attr("xda:dynamicArrayProperties", &attributes);
    }

    // Write the <cellMetadata> element.
    fn write_cell_metadata(&mut self) {
        let attributes = vec![("count", "1".to_string())];

        self.writer.xml_start_tag_attr("cellMetadata", &attributes);
        self.writer.xml_start_tag("bk");

        // Write the rc element.
        self.write_rc();

        self.writer.xml_end_tag("bk");
        self.writer.xml_end_tag("cellMetadata");
    }

    // Write the <rc> element.
    fn write_rc(&mut self) {
        let attributes = vec![("t", "1".to_string()), ("v", "0".to_string())];

        self.writer.xml_empty_tag_attr("rc", &attributes);
    }
}
