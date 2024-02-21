// relationship - A module for creating the Excel .rel relationship file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

mod tests;

use crate::xmlwriter::XMLWriter;

pub struct Relationship {
    pub(crate) writer: XMLWriter,
    relationships: Vec<(String, String, String)>,
    id_num: u16,
}

impl Relationship {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new struct to to track Excel shared strings between worksheets.
    pub(crate) fn new() -> Relationship {
        let writer = XMLWriter::new();

        Relationship {
            writer,
            relationships: vec![],
            id_num: 1,
        }
    }

    // Add container relationship to xlsx .rels xml files.
    pub(crate) fn add_document_relationship(
        &mut self,
        rel_type: &str,
        target: &str,
        target_mode: &str,
    ) {
        let document_schema = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        self.relationships.push((
            format!("{document_schema}/{rel_type}"),
            target.to_string(),
            target_mode.to_string(),
        ));
    }

    // Add container relationship to xlsx .rels xml files.
    pub(crate) fn add_package_relationship(&mut self, rel_type: &str, target: &str) {
        let package_schema = "http://schemas.openxmlformats.org/package/2006/relationships";

        self.relationships.push((
            format!("{package_schema}/{rel_type}"),
            target.to_string(),
            String::new(),
        ));
    }

    // Add container relationship to xlsx .rels xml files.
    pub(crate) fn add_office_relationship(
        &mut self,
        version: &str,
        rel_type: &str,
        target: &str,
        target_mode: &str,
    ) {
        let office_schema = "http://schemas.microsoft.com/office";

        self.relationships.push((
            format!("{office_schema}/{version}/relationships/{rel_type}"),
            target.to_string(),
            target_mode.to_string(),
        ));
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the Relationships element.
        self.write_relationships();

        // Close the Relationships tag.
        self.writer.xml_end_tag("Relationships");
    }

    // Write the <Relationships> element.
    fn write_relationships(&mut self) {
        let xmlns = "http://schemas.openxmlformats.org/package/2006/relationships";
        let attributes = [("xmlns", xmlns)];

        self.writer.xml_start_tag("Relationships", &attributes);

        for relationship in self.relationships.clone() {
            // Write the Relationship element.
            self.write_relationship(relationship);
        }
    }

    // Write the <Relationship> element.
    fn write_relationship(&mut self, relationship: (String, String, String)) {
        let r_id = format!("rId{}", self.id_num);
        let (rel_type, target, target_mode) = relationship;

        self.id_num += 1;

        let mut attributes = vec![("Id", r_id), ("Type", rel_type), ("Target", target)];

        if !target_mode.is_empty() {
            attributes.push(("TargetMode", target_mode));
        }

        self.writer.xml_empty_tag("Relationship", &attributes);
    }
}
