mod app;
mod content_types;
mod core;
mod relationship;
mod shared_strings;
mod shared_strings_table;
mod styles;
mod theme;
mod xmlwriter;

use crate::app::App;
use crate::content_types::ContentTypes;
use crate::core::Core;
use crate::relationship::Relationship;
use crate::shared_strings::SharedStrings;
use crate::shared_strings_table::SharedStringsTable;
use crate::styles::Styles;
use crate::theme::Theme;
use crate::xmlwriter::XMLWriter;
use tempfile::tempfile;

// Test function to excercise sub-modules.
pub fn assemble_all() {
    let tempfile = tempfile().unwrap();

    let mut writer = XMLWriter::new(&tempfile);
    let mut string_table = SharedStringsTable::new();
    let mut shared_strings = SharedStrings::new(&mut writer);
    string_table.get_shared_string_index("hello");
    shared_strings.assemble_xml_file(&string_table);

    let mut writer = XMLWriter::new(&tempfile);
    let mut app = App::new(&mut writer);
    app.add_heading_pair("Worksheets", 1);
    app.add_part_name("Sheet1");
    app.assemble_xml_file();

    let mut writer = XMLWriter::new(&tempfile);
    let mut theme = Theme::new(&mut writer);
    theme.assemble_xml_file();

    let mut writer = XMLWriter::new(&tempfile);
    let mut core = Core::new(&mut writer);
    core.assemble_xml_file();

    let mut writer = XMLWriter::new(&tempfile);
    let mut rels = Relationship::new(&mut writer);
    rels.add_document_relationship("/worksheet", "worksheets/sheet1.xml");
    rels.assemble_xml_file();

    let mut writer = XMLWriter::new(&tempfile);
    let mut content_types = ContentTypes::new(&mut writer);
    content_types.add_default("jpeg", "image/jpeg");
    content_types.add_worksheet_name("sheet1");
    content_types.add_share_strings();
    content_types.assemble_xml_file();

    let mut writer = XMLWriter::new(&tempfile);
    let mut styles = Styles::new(&mut writer);
    styles.assemble_xml_file();
}

#[cfg(test)]
mod tests {
    #[test]
    fn test_lib_compiles() {
        assert_eq!(true, true);
    }
}
