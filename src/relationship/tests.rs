// Relationship unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod relationship_tests {

    use crate::relationship::Relationship;
    use crate::test_functions::xml_to_vec;
    use crate::xmlwriter;
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

        let got = xmlwriter::cursor_to_str(&rels.writer);
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
