// Core unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod core_tests {

    use crate::core::Core;
    use crate::{test_functions::xml_to_vec, DocProperties};
    use crate::{xmlwriter, ExcelDateTime};

    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let date = ExcelDateTime::from_ymd(2010, 1, 1).unwrap();
        let properties = DocProperties::new()
            .set_author("A User")
            .set_creation_datetime(&date);

        let mut core = Core::new();
        core.properties = properties;

        core.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&core.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
              <dc:creator>A User</dc:creator>
              <cp:lastModifiedBy>A User</cp:lastModifiedBy>
              <dcterms:created xsi:type="dcterms:W3CDTF">2010-01-01T00:00:00Z</dcterms:created>
              <dcterms:modified xsi:type="dcterms:W3CDTF">2010-01-01T00:00:00Z</dcterms:modified>
            </cp:coreProperties>
            "#,
        );

        assert_eq!(expected, got);
    }
}
