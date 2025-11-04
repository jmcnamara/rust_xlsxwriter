// comment - A module for creating the Excel Comment.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::collections::BTreeMap;
use std::io::Cursor;

use crate::xmlwriter::{
    xml_data_element, xml_data_element_only, xml_declaration, xml_empty_tag, xml_empty_tag_only,
    xml_end_tag, xml_start_tag, xml_start_tag_only,
};
use crate::{utility, ColNum, Note, RowNum};

/// A struct to represent a Comment.
///
/// Comment is the older name for Note.
///
pub struct Comment {
    pub(crate) writer: Cursor<Vec<u8>>,
    pub(crate) notes: BTreeMap<RowNum, BTreeMap<ColNum, Note>>,
    pub(crate) note_authors: Vec<String>,
}

impl Comment {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    // Create a new Comment struct.
    pub(crate) fn new() -> Comment {
        let writer = Cursor::new(Vec::with_capacity(2048));

        Comment {
            writer,
            notes: BTreeMap::new(),
            note_authors: vec![],
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and generate the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        // Write the comments element.
        self.write_comments();

        // Write the authors element.
        self.write_authors();

        // Write the commentList element.
        self.write_comment_list();

        // Close the comments tag.
        xml_end_tag(&mut self.writer, "comments");
    }

    // Write the <comments> element.
    fn write_comments(&mut self) {
        let attributes = [(
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        )];

        xml_start_tag(&mut self.writer, "comments", &attributes);
    }

    // Write the <authors> element.
    fn write_authors(&mut self) {
        if self.note_authors.is_empty() {
            return;
        }

        xml_start_tag_only(&mut self.writer, "authors");

        for author in &self.note_authors.clone() {
            // Write the author element.
            self.write_author(author);
        }

        xml_end_tag(&mut self.writer, "authors");
    }

    // Write the <author> element.
    fn write_author(&mut self, author: &str) {
        xml_data_element_only(&mut self.writer, "author", author);
    }

    // Write the <commentList> element.
    fn write_comment_list(&mut self) {
        xml_start_tag_only(&mut self.writer, "commentList");

        for (row, columns) in &self.notes.clone() {
            for (col, note) in columns {
                // Write the comment element.
                self.write_comment(*row, *col, note);
            }
        }

        xml_end_tag(&mut self.writer, "commentList");
    }

    // Write the <comment> element.
    fn write_comment(&mut self, row: RowNum, col: ColNum, note: &Note) {
        let cell = utility::row_col_to_cell(row, col);
        let attributes = vec![("ref", cell), ("authorId", note.author_id.to_string())];

        xml_start_tag(&mut self.writer, "comment", &attributes);

        // Write the text element.
        self.write_text_block(note);

        xml_end_tag(&mut self.writer, "comment");
    }

    // Write the <text> element.
    fn write_text_block(&mut self, note: &Note) {
        xml_start_tag_only(&mut self.writer, "text");

        // Write the rPr element.
        if note.has_author_prefix {
            // Write the bold author run.
            xml_start_tag_only(&mut self.writer, "r");
            self.write_paragraph_run(note, true);

            let author = match &self.note_authors.get(note.author_id) {
                Some(author) => format!("{author}:"),
                None => "Author:".to_string(),
            };

            self.write_text(&author);
            xml_end_tag(&mut self.writer, "r");

            // Write the text on a new line.
            xml_start_tag_only(&mut self.writer, "r");
            self.write_paragraph_run(note, false);

            let text = format!("\n{}", note.text);
            self.write_text(&text);

            xml_end_tag(&mut self.writer, "r");
        } else {
            xml_start_tag_only(&mut self.writer, "r");
            self.write_paragraph_run(note, false);
            self.write_text(&note.text);
            xml_end_tag(&mut self.writer, "r");
        }

        xml_end_tag(&mut self.writer, "text");
    }

    // Write the <rPr> element.
    fn write_paragraph_run(&mut self, note: &Note, has_bold: bool) {
        xml_start_tag_only(&mut self.writer, "rPr");

        if has_bold {
            xml_empty_tag_only(&mut self.writer, "b");
        }

        // Write the sz element.
        self.write_font_size(note);

        // Write the color element.
        self.write_font_color();

        // Write the rFont element.
        self.write_font_name(note);

        // Write the family element.
        self.write_font_family(note);

        xml_end_tag(&mut self.writer, "rPr");
    }

    // Write the <sz> element.
    fn write_font_size(&mut self, note: &Note) {
        let attributes = [("val", note.format.font.size.clone())];

        xml_empty_tag(&mut self.writer, "sz", &attributes);
    }

    // Write the <color> element.
    fn write_font_color(&mut self) {
        let attributes = [("indexed", "81".to_string())];

        xml_empty_tag(&mut self.writer, "color", &attributes);
    }

    // Write the <rFont> element.
    fn write_font_name(&mut self, note: &Note) {
        let attributes = [("val", note.format.font.name.clone())];

        xml_empty_tag(&mut self.writer, "rFont", &attributes);
    }

    // Write the <family> element.
    fn write_font_family(&mut self, note: &Note) {
        let attributes = [("val", note.format.font.family.to_string())];

        xml_empty_tag(&mut self.writer, "family", &attributes);
    }

    // Write the <t> element.
    fn write_text(&mut self, text: &str) {
        let whitespace = ['\t', '\n', ' '];
        let attributes = if text.starts_with(whitespace) || text.ends_with(whitespace) {
            vec![("xml:space", "preserve")]
        } else {
            vec![]
        };

        xml_data_element(&mut self.writer, "t", text, &attributes);
    }
}
