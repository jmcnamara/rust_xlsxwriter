// Workbook unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod workbook_tests {

    use crate::{test_functions::xml_to_vec, XlsxError};
    use crate::{xmlwriter, Table, Workbook};
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let mut workbook = Workbook::default();
        workbook.add_worksheet();

        workbook.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&workbook.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
              <workbookPr defaultThemeVersion="124226"/>
              <bookViews>
                <workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>
              </bookViews>
              <sheets>
                <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
              </sheets>
              <calcPr calcId="124519" fullCalcOnLoad="1"/>
            </workbook>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn define_name() {
        let mut workbook = Workbook::default();

        // Test invalid defined names.
        let names = vec![
            ".foo",    // Invalid start character.
            "foo bar", // Space in name
            "Foo,",    // Other invalid characters.
            "Foo/", "Foo[", "Foo]", "Foo'", "Foo\"bar", "Foo:", "Foo*",
        ];

        for name in names {
            let result = workbook.define_name(name, "");
            assert!(matches!(result, Err(XlsxError::ParameterError(_))));
        }
    }

    #[test]
    fn duplicate_worksheets() {
        let mut workbook = Workbook::default();

        let _ = workbook.add_worksheet().set_name("Foo").unwrap();
        let _ = workbook.add_worksheet().set_name("Foo").unwrap();

        let result = workbook.save_to_buffer();
        assert!(matches!(result, Err(XlsxError::SheetnameReused(_))));
    }

    #[test]
    fn duplicate_worksheets_case_insensitive() {
        let mut workbook = Workbook::default();

        let _ = workbook.add_worksheet().set_name("Foo").unwrap();
        let _ = workbook.add_worksheet().set_name("foo").unwrap();

        let result = workbook.save_to_buffer();
        assert!(matches!(result, Err(XlsxError::SheetnameReused(_))));
    }

    #[test]
    fn duplicate_tables() {
        let mut workbook = Workbook::default();
        let worksheet = workbook.add_worksheet();

        let mut table = Table::new().set_name("Foo");

        worksheet.add_table(0, 0, 9, 9, &table).unwrap();

        table = table.set_name("foo");
        worksheet.add_table(10, 10, 19, 19, &table).unwrap();

        let result = workbook.prepare_tables();

        assert!(matches!(result, Err(XlsxError::TableNameReused(_))));
    }

    #[test]
    fn non_xml_theme() {
        let mut workbook = Workbook::default();
        let theme_file = "tests/input/themes/empty.xml";

        let result = workbook.use_custom_theme(theme_file);

        assert!(matches!(result, Err(XlsxError::ThemeError(_))));
    }

    #[test]
    fn image_gradient_fills_in_theme() {
        let mut workbook = Workbook::default();
        let theme_file = "tests/input/themes/civic.xml";

        let result = workbook.use_custom_theme(theme_file);

        assert!(matches!(result, Err(XlsxError::ThemeError(_))));
    }

    #[test]
    fn no_theme_file_in_zip() {
        let mut workbook = Workbook::default();
        let theme_file = "tests/input/themes/no_theme.zip";

        let result = workbook.use_custom_theme(theme_file);

        assert!(matches!(result, Err(XlsxError::ThemeError(_))));
    }
}
