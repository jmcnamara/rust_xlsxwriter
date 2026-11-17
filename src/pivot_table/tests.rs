// Pivot table unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod pivot_table_tests {

    use crate::pivot_table::parse_source_range;
    use crate::worksheet::{COL_MAX, ROW_MAX};
    use crate::{
        Format, Formula, PivotTable, PivotTableValueField, Table, Workbook, Worksheet, XlsxError,
    };

    #[test]
    fn test_validate_repeated_fields() {
        let mut worksheet = Worksheet::new();

        // Check that a field can't be used twice in the same axis.
        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 50, 3))
            .add_row_field("Region")
            .add_row_field("Region");

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check that a field can't be used in more than one of the row, column
        // and filter axes.
        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 50, 3))
            .add_row_field("Region")
            .add_column_field("Region");

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 50, 3))
            .add_column_field("Item")
            .add_filter_field("Item");

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check that a field can be used as a row field and a value field, and
        // in several value fields.
        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 50, 3))
            .add_row_field("Volume")
            .add_value_field(&PivotTableValueField::new("Volume"))
            .add_value_field(&PivotTableValueField::new("Volume"));

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_data_source() {
        let mut worksheet = Worksheet::new();

        // Check for a missing data source.
        let pivot_table = PivotTable::new().add_row_field("Region");

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check for a data source with a header row but no data rows.
        let pivot_table = PivotTable::new().set_data_source(("Sheet1", 0, 0, 0, 3));

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check for user defined field names that don't match the number of
        // source columns.
        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 50, 3))
            .set_field_names(&["Region", "Item", "Volume"]);

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check for a data source with the rows or columns reversed.
        let pivot_table = PivotTable::new().set_data_source(("Sheet1", 50, 0, 0, 3));

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::RowColumnOrderError)));

        let pivot_table = PivotTable::new().set_data_source(("Sheet1", 0, 3, 50, 0));

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::RowColumnOrderError)));

        // Check for a data source outside Excel's row or column limits.
        let pivot_table = PivotTable::new().set_data_source(("Sheet1", 0, 0, ROW_MAX, 3));

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let pivot_table = PivotTable::new().set_data_source(("Sheet1", 0, 0, 50, COL_MAX));

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));
    }

    #[test]
    fn test_validate_source_sheet() {
        // Check for a data source on a worksheet that doesn't exist.
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.write(0, 0, "Region").unwrap();
        worksheet.write(1, 0, "East").unwrap();

        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet3", 0, 0, 1, 0))
            .add_row_field("Region");
        let worksheet = workbook.add_worksheet();
        worksheet.add_pivot_table(2, 0, &pivot_table).unwrap();

        let result = workbook.save_to_buffer();
        assert!(matches!(
            result,
            Err(XlsxError::UnknownWorksheetNameOrIndex(_))
        ));
    }

    #[test]
    #[cfg(feature = "constant_memory")]
    fn test_validate_constant_memory_source() {
        // Check for a data source on a constant or low memory worksheet without
        // user defined field names.
        for low_memory in [false, true] {
            let mut workbook = Workbook::new();
            let worksheet = if low_memory {
                workbook.add_worksheet_with_low_memory().unwrap()
            } else {
                workbook.add_worksheet_with_constant_memory().unwrap()
            };
            worksheet.write(0, 0, "Region").unwrap();
            worksheet.write(1, 0, "East").unwrap();

            let pivot_table = PivotTable::new()
                .set_data_source(("Sheet1", 0, 0, 1, 0))
                .add_row_field("Region");
            let worksheet = workbook.add_worksheet();
            worksheet.add_pivot_table(2, 0, &pivot_table).unwrap();

            let result = workbook.save_to_buffer();
            assert!(matches!(result, Err(XlsxError::PivotTableError(_))));
        }

        // Check that user defined field names are allowed.
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet_with_constant_memory().unwrap();
        worksheet.write(0, 0, "Region").unwrap();
        worksheet.write(1, 0, "East").unwrap();

        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 1, 0))
            .set_field_names(&["Region"])
            .add_row_field("Region");
        let worksheet = workbook.add_worksheet();
        worksheet.add_pivot_table(2, 0, &pivot_table).unwrap();

        let result = workbook.save_to_buffer();
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_user_field_names() {
        // Check that user defined field names must match the header row on a
        // standard worksheet.
        for (field_names, is_ok) in [(["Region", "Volume"], true), (["Region", "Item"], false)] {
            let mut workbook = Workbook::new();
            let worksheet = workbook.add_worksheet();
            worksheet.write(0, 0, "Region").unwrap();
            worksheet.write(0, 1, "Volume").unwrap();
            worksheet.write(1, 0, "East").unwrap();
            worksheet.write(1, 1, 100).unwrap();

            let pivot_table = PivotTable::new()
                .set_data_source(("Sheet1", 0, 0, 1, 1))
                .set_field_names(&field_names)
                .add_row_field("Region");
            let worksheet = workbook.add_worksheet();
            worksheet.add_pivot_table(2, 0, &pivot_table).unwrap();

            let result = workbook.save_to_buffer();
            if is_ok {
                assert!(result.is_ok());
            } else {
                assert!(matches!(result, Err(XlsxError::PivotTableError(_))));
            }
        }
    }

    #[test]
    fn test_validate_filter_row_limit() {
        let mut worksheet = Worksheet::new();

        // Check for filter fields that move the pivot table body past the last
        // row. The body is below the filters and a blank row.
        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 50, 3))
            .add_row_field("Region")
            .add_filter_field("Month");

        let result = worksheet.add_pivot_table(ROW_MAX - 2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check that the body can start in the last row.
        let result = worksheet.add_pivot_table(ROW_MAX - 3, 0, &pivot_table);
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_overlap() {
        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 50, 3))
            .add_row_field("Region");

        // Check for a pivot table on a worksheet table.
        let mut worksheet = Worksheet::new();
        worksheet.add_table(2, 0, 8, 3, &Table::new()).unwrap();

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check for a worksheet table on a pivot table.
        let mut worksheet = Worksheet::new();
        worksheet.add_pivot_table(2, 0, &pivot_table).unwrap();

        let result = worksheet.add_table(2, 0, 8, 3, &Table::new());
        assert!(matches!(result, Err(XlsxError::TableError(_))));

        // Check for a pivot table on a merged range.
        let mut worksheet = Worksheet::new();
        worksheet
            .merge_range(2, 0, 2, 3, "Merged", &Format::new())
            .unwrap();

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check for two pivot tables at the same position.
        let mut worksheet = Worksheet::new();
        worksheet.add_pivot_table(2, 0, &pivot_table).unwrap();

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check for a pivot table whose filter fields overlap a previous pivot
        // table. The body is below the filters and a blank row.
        let mut worksheet = Worksheet::new();
        worksheet.add_pivot_table(4, 0, &pivot_table).unwrap();

        let filtered_pivot_table = pivot_table.clone().add_filter_field("Month");
        let result = worksheet.add_pivot_table(2, 0, &filtered_pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check for a pivot table on its own source data.
        let mut worksheet = Worksheet::new();
        worksheet.set_name("Sheet1").unwrap();

        let result = worksheet.add_pivot_table(10, 1, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check that pivot tables and tables apart from each other are allowed.
        let mut worksheet = Worksheet::new();
        worksheet.add_table(0, 6, 4, 7, &Table::new()).unwrap();
        worksheet.add_pivot_table(2, 0, &pivot_table).unwrap();

        let result = worksheet.add_pivot_table(2, 3, &pivot_table);
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_chartsheet() {
        // Check for a pivot table added to a chartsheet.
        let mut workbook = Workbook::new();
        let chartsheet = workbook.add_chartsheet();

        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 50, 3))
            .add_row_field("Region");

        let result = chartsheet.add_pivot_table(0, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));
    }

    #[test]
    fn test_validate_field_names() {
        // Check for a field name that doesn't match a source header in the
        // row, column, filter or value fields.
        let pivot_tables = [
            PivotTable::new().add_row_field("Legion"),
            PivotTable::new().add_column_field("Legion"),
            PivotTable::new().add_filter_field("Legion"),
            PivotTable::new().add_value_field(&PivotTableValueField::new("Legion")),
        ];

        for pivot_table in pivot_tables {
            let mut workbook = Workbook::new();
            let worksheet = workbook.add_worksheet();
            worksheet.write(0, 0, "Region").unwrap();
            worksheet.write(0, 1, "Volume").unwrap();
            worksheet.write(1, 0, "East").unwrap();
            worksheet.write(1, 1, 100).unwrap();

            let pivot_table = pivot_table.set_data_source(("Sheet1", 0, 0, 1, 1));
            let worksheet = workbook.add_worksheet();
            worksheet.add_pivot_table(2, 0, &pivot_table).unwrap();

            let result = workbook.save_to_buffer();
            assert!(matches!(result, Err(XlsxError::PivotTableError(_))));
        }
    }

    #[test]
    fn test_validate_value_field_names() {
        let mut worksheet = Worksheet::new();

        // Check for two value fields with the same custom name.
        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 50, 3))
            .add_value_field(&PivotTableValueField::new("Volume").set_custom_name("Total"))
            .add_value_field(&PivotTableValueField::new("Volume").set_custom_name("Total"));

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check for two value fields with custom names that differ only by case.
        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 50, 3))
            .add_value_field(&PivotTableValueField::new("Volume").set_custom_name("Total"))
            .add_value_field(&PivotTableValueField::new("Volume").set_custom_name("TOTAL"));

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check for a custom name that is the same as a default name.
        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 50, 3))
            .add_value_field(&PivotTableValueField::new("Volume"))
            .add_value_field(&PivotTableValueField::new("Volume").set_custom_name("Sum of Volume"));

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check that repeated default names are allowed since they are made
        // unique.
        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 50, 3))
            .add_value_field(&PivotTableValueField::new("Volume"))
            .add_value_field(&PivotTableValueField::new("Volume"));

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(result.is_ok());

        // Check for a custom name that is the same as a source header, ignoring case.
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.write(0, 0, "Region").unwrap();
        worksheet.write(0, 1, "Volume").unwrap();
        worksheet.write(1, 0, "East").unwrap();
        worksheet.write(1, 1, 100).unwrap();

        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 1, 1))
            .add_row_field("Region")
            .add_value_field(&PivotTableValueField::new("Volume").set_custom_name("VOLUME"));
        let worksheet = workbook.add_worksheet();
        worksheet.add_pivot_table(2, 0, &pivot_table).unwrap();

        let result = workbook.save_to_buffer();
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));
    }

    #[test]
    fn test_header_cell_types() {
        // Check that formula, boolean, rich string and number header cells can
        // be used as field names.
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        let bold = Format::new().set_bold();
        worksheet
            .write_formula(0, 0, Formula::new("=\"Reg\"&\"ion\"").set_result("Region"))
            .unwrap();
        worksheet.write(0, 1, true).unwrap();
        worksheet
            .write_rich_string(0, 2, &[(&bold, "Vol"), (&Format::default(), "ume")])
            .unwrap();
        worksheet.write(0, 3, 2024).unwrap();
        worksheet.write(1, 0, "East").unwrap();

        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 1, 3))
            .add_row_field("Region")
            .add_column_field("TRUE")
            .add_filter_field("2024")
            .add_value_field(&PivotTableValueField::new("Volume"));
        let worksheet = workbook.add_worksheet();
        worksheet.add_pivot_table(2, 0, &pivot_table).unwrap();

        let result = workbook.save_to_buffer();
        assert!(result.is_ok());

        // Check for a formula header cell without a result.
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet
            .write_formula(0, 0, Formula::new("=\"Reg\"&\"ion\""))
            .unwrap();
        worksheet.write(1, 0, "East").unwrap();

        let pivot_table = PivotTable::new()
            .set_data_source(("Sheet1", 0, 0, 1, 0))
            .add_row_field("Region");
        let worksheet = workbook.add_worksheet();
        worksheet.add_pivot_table(2, 0, &pivot_table).unwrap();

        let result = workbook.save_to_buffer();
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));
    }

    #[test]
    fn test_validate_headers() {
        // Check for a blank header cell, a missing header cell, a missing
        // header row or a duplicate header, ignoring case.
        let headers = [
            vec!["Region", "", "Volume"],
            vec!["Region", "Volume"],
            vec![],
            vec!["Region", "Item", "Region"],
            vec!["Region", "Item", "REGION"],
        ];

        for header in headers {
            let mut workbook = Workbook::new();
            let worksheet = workbook.add_worksheet();
            for (col, name) in header.iter().enumerate() {
                worksheet.write(0, col as u16, *name).unwrap();
            }
            worksheet.write(1, 0, "East").unwrap();

            let pivot_table = PivotTable::new()
                .set_data_source(("Sheet1", 0, 0, 1, 2))
                .add_row_field("Region");
            let worksheet = workbook.add_worksheet();
            worksheet.add_pivot_table(2, 0, &pivot_table).unwrap();

            let result = workbook.save_to_buffer();
            assert!(matches!(result, Err(XlsxError::PivotTableError(_))));
        }
    }

    #[test]
    fn test_parse_source_range() {
        // Check for valid range strings.
        let ranges = [
            ("Sheet1!A1:D9", "Sheet1", 0, 0, 8, 3),
            ("Sheet1!$A$1:$D$9", "Sheet1", 0, 0, 8, 3),
            ("=Sheet1!A1:D9", "Sheet1", 0, 0, 8, 3),
            ("Sheet1!a1:d9", "Sheet1", 0, 0, 8, 3),
            ("'Sales Data'!B3:AA100", "Sales Data", 2, 1, 99, 26),
            ("'Bob''s Data'!A1:B2", "Bob's Data", 0, 0, 1, 1),
            ("Sheet1!A1:XFD1048576", "Sheet1", 0, 0, 1_048_575, 16_383),
        ];

        for (range, sheet_name, first_row, first_col, last_row, last_col) in ranges {
            assert_eq!(
                parse_source_range(range),
                Some((
                    sheet_name.to_string(),
                    first_row,
                    first_col,
                    last_row,
                    last_col
                )),
                "{range}"
            );
        }

        // Check for invalid range strings.
        let ranges = [
            "",
            "A1:D9",
            "!A1:D9",
            "Sheet1!",
            "Sheet1!A1",
            "Sheet1!A:D",
            "Sheet1!A0:D9",
            "Sheet1!A1:D9:E10",
            "Sheet1!AAAA1:D9",
            "'Sheet 1!A1:D9",
            "Sheet1!Table1",
        ];

        for range in ranges {
            assert_eq!(parse_source_range(range), None, "{range}");
        }
    }

    #[test]
    fn test_validate_source_range_string() {
        let mut worksheet = Worksheet::new();

        // Check for an invalid data source range string.
        let pivot_table = PivotTable::new()
            .set_data_source("Sheet1!A1")
            .add_row_field("Region");

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        // Check that a valid range string replaces an earlier invalid one.
        let pivot_table = PivotTable::new()
            .set_data_source("Sheet1!A1")
            .set_data_source("Sheet1!A1:D9")
            .add_row_field("Region");

        let result = worksheet.add_pivot_table(2, 0, &pivot_table);
        assert!(result.is_ok());

        // Check for a column beyond the Excel limit.
        let pivot_table = PivotTable::new().set_data_source("Sheet1!A1:XFE9");

        let result = worksheet.add_pivot_table(20, 0, &pivot_table);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));
    }
}
