// Comment unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod comment_tests {

    use crate::comment::Comment;
    use crate::test_functions::xml_to_vec;
    use crate::{xmlwriter, Note};
    use pretty_assertions::assert_eq;
    use std::collections::BTreeMap;

    #[test]
    fn test_note_on_chartsheet() {
        // Check for a note added to a chartsheet.
        let mut chartsheet = crate::Worksheet::new_chartsheet();

        let result = chartsheet.insert_note(0, 0, &Note::new("Note"));
        assert!(matches!(result, Err(crate::XlsxError::ParameterError(_))));
    }

    #[test]
    fn test_assemble_with_escapes() {
        let mut comment = Comment::new();

        let note = Note::new("a_x0041_b").add_author_prefix(false);
        comment.notes = BTreeMap::from([(0, BTreeMap::from([(0, note)]))]);
        comment.note_authors = vec!["a_x0041_b".to_string()];

        comment.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&comment.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <authors>
                <author>a_x005F_x0041_b</author>
              </authors>
              <commentList>
                <comment ref="A1" authorId="0">
                  <text>
                    <r>
                      <rPr>
                        <sz val="8"/>
                        <color indexed="81"/>
                        <rFont val="Tahoma"/>
                        <family val="2"/>
                      </rPr>
                      <t>a_x005F_x0041_b</t>
                    </r>
                  </text>
                </comment>
              </commentList>
            </comments>
            "#,
        );

        assert_eq!(expected, got);
    }
}
