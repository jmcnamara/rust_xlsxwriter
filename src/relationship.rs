// relationship - A module for creating the Excel .rel relationship file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

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
            "".to_string(),
        ));
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
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

        self.writer.xml_start_tag_attr("Relationships", &attributes);

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

        self.writer.xml_empty_tag_attr("Relationship", &attributes);
    }
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use crate::relationship::Relationship;
    use crate::test_functions::xml_to_vec;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let mut rels = Relationship::new();

        rels.add_document_relationship("worksheet", "worksheets/sheet1.xml", "");
        rels.add_document_relationship("theme", "theme/theme1.xml", "");
        rels.add_document_relationship("styles", "styles.xml", "");
        rels.add_document_relationship("sharedStrings", "sharedStrings.xml", "");
        rels.add_document_relationship("calcChain", "calcChain.xml", "");

        rels.assemble_xml_file();

        let got = rels.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
              <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
              <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
              <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
              <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
            </Relationships>
            "#,
        );

        assert_eq!(expected, got);
    }
}
