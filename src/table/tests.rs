// Table unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod table_tests {

    use crate::table::Table;
    use crate::test_functions::xml_to_vec;
    use crate::{TableColumn, TableFunction, Worksheet, XlsxError};
    use pretty_assertions::assert_eq;

    #[test]
    fn test_row_methods() {
        let mut table = Table::new();
        table.cell_range.first_row = 0;
        table.cell_range.first_col = 0;
        table.cell_range.last_row = 9;
        table.cell_range.last_col = 4;

        assert_eq!(1, table.first_data_row());
        assert_eq!(9, table.last_data_row());

        table = table.set_total_row(true);
        assert_eq!(1, table.first_data_row());
        assert_eq!(8, table.last_data_row());

        table = table.set_header_row(false);
        assert_eq!(0, table.first_data_row());
        assert_eq!(8, table.last_data_row());
    }

    #[test]
    fn test_column_validation() {
        // Test the table column validation and checks.
        let mut table = Table::new();
        let default_headers = vec![
            String::from("Column1"),
            String::from("Column2"),
            String::from("Column3"),
            String::from("Column4"),
        ];

        table.cell_range.first_row = 0;
        table.cell_range.first_col = 0;
        table.cell_range.last_row = 8;
        table.cell_range.last_col = 4;
        table.index = 1;

        let columns = vec![
            TableColumn::new().set_header("Foo"),
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::new().set_header("foo"), // Lowercase duplicate.
        ];

        table = table.set_columns(&columns);
        let result = table.initialize_columns(&default_headers);

        assert!(matches!(result, Err(XlsxError::TableError(_))));

        let columns = vec![
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::new().set_header("column1"), // Lowercase duplicate.
        ];

        table = table.set_columns(&columns);
        let result = table.initialize_columns(&default_headers);

        assert!(matches!(result, Err(XlsxError::TableError(_))));
    }

    #[test]
    fn test_assemble1() {
        let mut table = Table::new();
        let default_headers = vec![
            String::from("Column1"),
            String::from("Column2"),
            String::from("Column3"),
            String::from("Column4"),
        ];

        table.cell_range.first_row = 2;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 12;
        table.cell_range.last_col = 5;
        table.index = 1;

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F13" totalsRowShown="0">
                <autoFilter ref="C3:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble2() {
        let mut table = Table::new().set_style(crate::TableStyle::Light17);
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 3;
        table.cell_range.first_col = 3;
        table.cell_range.last_row = 14;
        table.cell_range.last_col = 8;
        table.index = 2;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="2" name="Table2" displayName="Table2" ref="D4:I15" totalsRowShown="0">
                <autoFilter ref="D4:I15"/>
                <tableColumns count="6">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                    <tableColumn id="5" name="Column5"/>
                    <tableColumn id="6" name="Column6"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleLight17" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble3() {
        let mut table = Table::new()
            .set_first_column(true)
            .set_last_column(true)
            .set_banded_rows(false)
            .set_banded_columns(true);

        let worksheet = Worksheet::new();

        table.cell_range.first_row = 4;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 15;
        table.cell_range.last_col = 3;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C5:D16" totalsRowShown="0">
                <autoFilter ref="C5:D16"/>
                <tableColumns count="2">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="1" showLastColumn="1" showRowStripes="0" showColumnStripes="1"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble4() {
        let mut table = Table::new().set_autofilter(false);
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 2;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 12;
        table.cell_range.last_col = 5;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F13" totalsRowShown="0">
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble5() {
        let mut table = Table::new().set_header_row(false);
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 3;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 12;
        table.cell_range.last_col = 5;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C4:F13" headerRowCount="0" totalsRowShown="0">
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble6() {
        let mut table = Table::new();
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 2;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 12;
        table.cell_range.last_col = 5;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        let columns = vec![
            TableColumn::new().set_header("Foo"),
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::new().set_header("Baz"),
        ];

        table = table.set_columns(&columns);

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F13" totalsRowShown="0">
                <autoFilter ref="C3:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Foo"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Baz"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);

        // Try a variation with too many columns settings. It should be
        // truncated to the actual number of columns.
        let mut table = Table::new();

        table.cell_range.first_row = 2;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 12;
        table.cell_range.last_col = 5;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        let columns = vec![
            TableColumn::new().set_header("Foo"),
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::new().set_header("Baz"),
            TableColumn::new().set_header("Too many"),
        ];

        table = table.set_columns(&columns);

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble7() {
        let mut table = Table::new().set_total_row(true);
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 2;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 13;
        table.cell_range.last_col = 5;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F14" totalsRowCount="1">
                <autoFilter ref="C3:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble8() {
        let mut table = Table::new();
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 2;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 13;
        table.cell_range.last_col = 5;
        table.index = 1;

        let columns = vec![
            TableColumn::new().set_total_label("Total"),
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::new().set_total_function(TableFunction::Count),
        ];

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        table = table.set_columns(&columns).set_total_row(true);

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F14" totalsRowCount="1">
                <autoFilter ref="C3:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1" totalsRowLabel="Total"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4" totalsRowFunction="count"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble9() {
        let mut table = Table::new();
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 1;
        table.cell_range.first_col = 1;
        table.cell_range.last_row = 7;
        table.cell_range.last_col = 10;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        let columns = vec![
            TableColumn::new()
                .set_total_label("Total")
                .set_total_function(TableFunction::Max), // Max should be ignore due to label.
            TableColumn::new().set_total_function(TableFunction::None), // Should be ignored.
            TableColumn::new().set_total_function(TableFunction::Average),
            TableColumn::new().set_total_function(TableFunction::Count),
            TableColumn::new().set_total_function(TableFunction::CountNumbers),
            TableColumn::new().set_total_function(TableFunction::Max),
            TableColumn::new().set_total_function(TableFunction::Min),
            TableColumn::new().set_total_function(TableFunction::Sum),
            TableColumn::new().set_total_function(TableFunction::StdDev),
            TableColumn::new().set_total_function(TableFunction::Var),
        ];

        table = table.set_columns(&columns).set_total_row(true);

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="B2:K8" totalsRowCount="1">
                <autoFilter ref="B2:K7"/>
                <tableColumns count="10">
                    <tableColumn id="1" name="Column1" totalsRowLabel="Total"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3" totalsRowFunction="average"/>
                    <tableColumn id="4" name="Column4" totalsRowFunction="count"/>
                    <tableColumn id="5" name="Column5" totalsRowFunction="countNums"/>
                    <tableColumn id="6" name="Column6" totalsRowFunction="max"/>
                    <tableColumn id="7" name="Column7" totalsRowFunction="min"/>
                    <tableColumn id="8" name="Column8" totalsRowFunction="sum"/>
                    <tableColumn id="9" name="Column9" totalsRowFunction="stdDev"/>
                    <tableColumn id="10" name="Column10" totalsRowFunction="var"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble10() {
        let mut table = Table::new().set_name("MyTable");
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 1;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 12;
        table.cell_range.last_col = 5;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="MyTable" displayName="MyTable" ref="C2:F13" totalsRowShown="0">
                <autoFilter ref="C2:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }
}
