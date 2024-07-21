// comment - A module for creating the Excel Comment.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::collections::BTreeMap;

use crate::{utility, xmlwriter::XMLWriter, ColNum, Note, RowNum};

/// A struct to represent a Comment.
///
/// TODO.
pub struct Comment {
    pub(crate) writer: XMLWriter,
    pub(crate) notes: BTreeMap<RowNum, BTreeMap<ColNum, Note>>,
    pub(crate) note_authors: BTreeMap<String, usize>,
}

impl Comment {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    // Create a new Comment struct.
    pub(crate) fn new() -> Comment {
        let writer = XMLWriter::new();

        Comment {
            writer,
            notes: BTreeMap::new(),
            note_authors: BTreeMap::new(),
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the comments element.
        self.write_comments();

        // Write the authors element.
        self.write_authors();

        // Write the commentList element.
        self.write_comment_list();

        // Close the comments tag.
        self.writer.xml_end_tag("comments");
    }

    // Write the <comments> element.
    fn write_comments(&mut self) {
        let attributes = [(
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        )];

        self.writer.xml_start_tag("comments", &attributes);
    }

    // Write the <authors> element.
    fn write_authors(&mut self) {
        if self.note_authors.is_empty() {
            return;
        }

        self.writer.xml_start_tag_only("authors");

        let authors: Vec<String> = self.note_authors.keys().cloned().collect();
        for author in &authors {
            // Write the author element.
            self.write_author(author);
        }

        self.writer.xml_end_tag("authors");
    }

    // Write the <author> element.
    fn write_author(&mut self, author: &str) {
        self.writer.xml_data_element_only("author", author);
    }

    // Write the <commentList> element.
    fn write_comment_list(&mut self) {
        self.writer.xml_start_tag_only("commentList");

        for (row, columns) in &self.notes.clone() {
            for (col, note) in columns {
                // Write the comment element.
                self.write_comment(*row, *col, note);
            }
        }

        self.writer.xml_end_tag("commentList");
    }

    // Write the <comment> element.
    fn write_comment(&mut self, row: RowNum, col: ColNum, note: &Note) {
        let cell = utility::row_col_to_cell(row, col);
        let mut attributes = vec![("ref", cell)];

        if let Some(id) = note.author_id {
            attributes.push(("authorId", id.to_string()));
        }

        self.writer.xml_start_tag("comment", &attributes);

        // Write the text element.
        self.write_text_block(&note.text);

        self.writer.xml_end_tag("comment");
    }

    // Write the <text> element.
    fn write_text_block(&mut self, text: &str) {
        self.writer.xml_start_tag_only("text");
        self.writer.xml_start_tag_only("r");

        // Write the rPr element.
        self.write_paragraph_run();

        // Write the t text element.
        self.write_text(text);

        self.writer.xml_end_tag("r");
        self.writer.xml_end_tag("text");
    }

    // Write the <rPr> element.
    fn write_paragraph_run(&mut self) {
        self.writer.xml_start_tag_only("rPr");

        // Write the sz element.
        self.write_font_size();

        // Write the color element.
        self.write_font_color();

        // Write the rFont element.
        self.write_font_name();

        // Write the family element.
        self.write_font_family();

        self.writer.xml_end_tag("rPr");
    }

    // Write the <sz> element.
    fn write_font_size(&mut self) {
        let attributes = [("val", "8".to_string())];

        self.writer.xml_empty_tag("sz", &attributes);
    }

    // Write the <color> element.
    fn write_font_color(&mut self) {
        let attributes = [("indexed", "81".to_string())];

        self.writer.xml_empty_tag("color", &attributes);
    }

    // Write the <rFont> element.
    fn write_font_name(&mut self) {
        let attributes = [("val", "Tahoma".to_string())];

        self.writer.xml_empty_tag("rFont", &attributes);
    }

    // Write the <family> element.
    fn write_font_family(&mut self) {
        let attributes = [("val", "2".to_string())];

        self.writer.xml_empty_tag("family", &attributes);
    }

    // Write the <t> element.
    fn write_text(&mut self, text: &str) {
        self.writer.xml_data_element_only("t", text);
    }
}
