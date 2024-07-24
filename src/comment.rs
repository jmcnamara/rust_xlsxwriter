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
/// Comments is the older name for Notes.
///
pub struct Comment {
    pub(crate) writer: XMLWriter,
    pub(crate) notes: BTreeMap<RowNum, BTreeMap<ColNum, Note>>,
    pub(crate) note_authors: Vec<String>,
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
            note_authors: vec![],
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

        //let authors: Vec<String> = self.note_authors.keys().cloned().collect();
        for author in &self.note_authors.clone() {
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
        let attributes = vec![("ref", cell), ("authorId", note.author_id.to_string())];

        self.writer.xml_start_tag("comment", &attributes);

        // Write the text element.
        self.write_text_block(note);

        self.writer.xml_end_tag("comment");
    }

    // Write the <text> element.
    fn write_text_block(&mut self, note: &Note) {
        self.writer.xml_start_tag_only("text");

        // Write the rPr element.
        if note.has_author_prefix {
            // Write the bold author run.
            self.writer.xml_start_tag_only("r");
            self.write_paragraph_run(note, true);

            let author = match &self.note_authors.get(note.author_id) {
                Some(author) => format!("{author}:"),
                None => "Author:".to_string(),
            };

            self.write_text(&author);
            self.writer.xml_end_tag("r");

            // Write the text on a new line.
            self.writer.xml_start_tag_only("r");
            self.write_paragraph_run(note, false);

            let text = format!("\n{}", note.text);
            self.write_text(&text);

            self.writer.xml_end_tag("r");
        } else {
            self.writer.xml_start_tag_only("r");
            self.write_paragraph_run(note, false);
            self.write_text(&note.text);
            self.writer.xml_end_tag("r");
        }

        self.writer.xml_end_tag("text");
    }

    // Write the <rPr> element.
    fn write_paragraph_run(&mut self, note: &Note, has_bold: bool) {
        self.writer.xml_start_tag_only("rPr");

        if has_bold {
            self.writer.xml_empty_tag_only("b");
        }

        // Write the sz element.
        self.write_font_size(note);

        // Write the color element.
        self.write_font_color();

        // Write the rFont element.
        self.write_font_name(note);

        // Write the family element.
        self.write_font_family(note);

        self.writer.xml_end_tag("rPr");
    }

    // Write the <sz> element.
    fn write_font_size(&mut self, note: &Note) {
        let attributes = [("val", note.format.font.size.to_string())];

        self.writer.xml_empty_tag("sz", &attributes);
    }

    // Write the <color> element.
    fn write_font_color(&mut self) {
        let attributes = [("indexed", "81".to_string())];

        self.writer.xml_empty_tag("color", &attributes);
    }

    // Write the <rFont> element.
    fn write_font_name(&mut self, note: &Note) {
        let attributes = [("val", note.format.font.name.clone())];

        self.writer.xml_empty_tag("rFont", &attributes);
    }

    // Write the <family> element.
    fn write_font_family(&mut self, note: &Note) {
        let attributes = [("val", note.format.font.family.to_string())];

        self.writer.xml_empty_tag("family", &attributes);
    }

    // Write the <t> element.
    fn write_text(&mut self, text: &str) {
        let whitespace = ['\t', '\n', ' '];
        let attributes = if text.starts_with(whitespace) || text.ends_with(whitespace) {
            vec![("xml:space", "preserve")]
        } else {
            vec![]
        };

        self.writer.xml_data_element("t", text, &attributes);
    }
}
