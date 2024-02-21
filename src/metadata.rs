// metadata - A module for creating the Excel Metadata.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

use crate::xmlwriter::XMLWriter;

pub struct Metadata {
    pub(crate) writer: XMLWriter,
    pub(crate) has_dynamic_functions: bool,
    pub(crate) has_embedded_images: bool,
    pub(crate) num_embedded_images: u32,
}

impl Metadata {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new Metadata struct.
    pub fn new() -> Metadata {
        let writer = XMLWriter::new();

        Metadata {
            writer,
            has_dynamic_functions: false,
            has_embedded_images: false,
            num_embedded_images: 0,
        }
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
        if self.has_dynamic_functions {
            self.write_cell_future_metadata();
        }
        if self.has_embedded_images {
            self.write_value_future_metadata();
        }

        // Write the cellMetadata element.
        if self.has_dynamic_functions {
            self.write_cell_metadata();
        }
        if self.has_embedded_images {
            self.write_value_metadata();
        }

        // Close the metadata tag.
        self.writer.xml_end_tag("metadata");
    }

    // Write the <metadata> element.
    fn write_metadata(&mut self) {
        let mut attributes = vec![(
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        )];

        if self.has_embedded_images {
            attributes.push((
                "xmlns:xlrd",
                "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata",
            ));
        }

        if self.has_dynamic_functions {
            attributes.push((
                "xmlns:xda",
                "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray",
            ));
        }

        self.writer.xml_start_tag("metadata", &attributes);
    }

    // Write the <metadataTypes> element.
    fn write_metadata_types(&mut self) {
        let mut count = 0;

        if self.has_dynamic_functions {
            count += 1;
        }

        if self.has_embedded_images {
            count += 1;
        }

        let attributes = [("count", count.to_string())];

        self.writer.xml_start_tag("metadataTypes", &attributes);

        // Write the metadataType element.
        if self.has_dynamic_functions {
            self.write_cell_metadata_type();
        }
        if self.has_embedded_images {
            self.write_value_metadata_type();
        }

        self.writer.xml_end_tag("metadataTypes");
    }

    // Write the cell <metadataType> element.
    fn write_cell_metadata_type(&mut self) {
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

    // Write the value <metadataType> element.
    fn write_value_metadata_type(&mut self) {
        let attributes = [
            ("name", "XLRICHVALUE"),
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
        ];

        self.writer.xml_empty_tag("metadataType", &attributes);
    }

    // Write the cell <futureMetadata> element.
    fn write_cell_future_metadata(&mut self) {
        let attributes = [("name", "XLDAPR"), ("count", "1")];

        self.writer.xml_start_tag("futureMetadata", &attributes);
        self.writer.xml_start_tag_only("bk");
        self.writer.xml_start_tag_only("extLst");

        // Write the ext element.
        self.write_cell_ext();

        self.writer.xml_end_tag("extLst");
        self.writer.xml_end_tag("bk");
        self.writer.xml_end_tag("futureMetadata");
    }

    // Write the value <futureMetadata> element.
    fn write_value_future_metadata(&mut self) {
        let attributes = [
            ("name", "XLRICHVALUE".to_string()),
            ("count", self.num_embedded_images.to_string()),
        ];

        self.writer.xml_start_tag("futureMetadata", &attributes);

        // Write the ext element.
        for index in 0..self.num_embedded_images {
            self.writer.xml_start_tag_only("bk");
            self.writer.xml_start_tag_only("extLst");
            self.write_value_ext(index);
            self.writer.xml_end_tag("extLst");
            self.writer.xml_end_tag("bk");
        }

        self.writer.xml_end_tag("futureMetadata");
    }

    // Write the <ext> element.
    fn write_cell_ext(&mut self) {
        let attributes = [("uri", "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}")];

        self.writer.xml_start_tag("ext", &attributes);

        // Write the xda:dynamicArrayProperties element.
        self.write_xda_dynamic_array_properties();

        self.writer.xml_end_tag("ext");
    }

    // Write the <ext> element.
    fn write_value_ext(&mut self, index: u32) {
        let attributes = [("uri", "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}")];

        self.writer.xml_start_tag("ext", &attributes);

        // Write the xlrd:rvb element.
        self.write_xlrd_rvb(index);

        self.writer.xml_end_tag("ext");
    }

    // Write the <xlrd:rvb> element.
    fn write_xlrd_rvb(&mut self, index: u32) {
        let attributes = [("i", index.to_string())];

        self.writer.xml_empty_tag("xlrd:rvb", &attributes);
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
        self.write_rc(1, 0);

        self.writer.xml_end_tag("bk");
        self.writer.xml_end_tag("cellMetadata");
    }

    // Write the <valueMetadata> element.
    fn write_value_metadata(&mut self) {
        let attributes = [("count", self.num_embedded_images.to_string())];
        let rc_type = if self.has_dynamic_functions { 2 } else { 1 };

        self.writer.xml_start_tag("valueMetadata", &attributes);

        for index in 0..self.num_embedded_images {
            self.writer.xml_start_tag_only("bk");
            self.write_rc(rc_type, index);
            self.writer.xml_end_tag("bk");
        }

        self.writer.xml_end_tag("valueMetadata");
    }

    // Write the <rc> element.
    fn write_rc(&mut self, rc_type: u32, value: u32) {
        let attributes = [("t", rc_type.to_string()), ("v", value.to_string())];

        self.writer.xml_empty_tag("rc", &attributes);
    }
}
