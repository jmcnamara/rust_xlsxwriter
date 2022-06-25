// Entry point for rust_xlsxwriter library.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

mod app;
mod content_types;
mod core;
mod packager;
mod relationship;
mod shared_strings;
mod shared_strings_table;
mod styles;
mod test_functions;
mod theme;
mod workbook;
mod worksheet;
mod xmlwriter;

use crate::app::App;
use crate::content_types::ContentTypes;
use crate::core::Core;
use crate::relationship::Relationship;
use crate::shared_strings::SharedStrings;
use crate::shared_strings_table::SharedStringsTable;
use crate::styles::Styles;
use crate::theme::Theme;
//use crate::workbook::Workbook;
use crate::worksheet::Worksheet;

pub use workbook::*;

// Test function to excercise sub-modules.
pub fn assemble_all() {
    let mut string_table = SharedStringsTable::new();
    let mut shared_strings = SharedStrings::new();
    string_table.get_shared_string_index("hello");
    shared_strings.assemble_xml_file(&string_table);

    let mut app = App::new();
    app.add_heading_pair("Worksheets", 1);
    app.add_part_name("Sheet1");
    app.assemble_xml_file();

    let mut theme = Theme::new();
    theme.assemble_xml_file();

    let mut core = Core::new();
    core.assemble_xml_file();

    let mut rels = Relationship::new();
    rels.add_document_relationship("/worksheet", "worksheets/sheet1.xml");
    rels.assemble_xml_file();

    let mut content_types = ContentTypes::new();
    content_types.add_default("jpeg", "image/jpeg");
    content_types.add_worksheet_name("sheet1");
    content_types.add_share_strings();
    content_types.assemble_xml_file();

    let mut styles = Styles::new();
    styles.assemble_xml_file();

    let mut workbook = Workbook::new("test.xlsx");
    workbook.assemble_xml_file();
    workbook.close();

    let mut worksheet = Worksheet::new();
    worksheet.assemble_xml_file();
}

#[cfg(test)]
mod tests {
    #[test]
    fn test_lib_compiles() {
        assert_eq!(true, true);
    }
}
