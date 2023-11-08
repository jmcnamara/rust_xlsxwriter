// conditional_format unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod conditional_format_tests {

    use crate::test_functions::xml_to_vec;
    use crate::worksheet::*;
    use crate::ConditionalFormatCell;
    use crate::ConditionalFormatCellCriteria;
    use crate::ExcelDateTime;
    use crate::Formula;
    use crate::XlsxError;
    use pretty_assertions::assert_eq;

    #[test]
    fn quoted_string_01() {
        let conditional_format = ConditionalFormatCell::new()
            .set_criteria(ConditionalFormatCellCriteria::EqualTo)
            .set_value(5);

        let got = conditional_format.get_rule_string(None, 1);
        let expected =
            r#"<cfRule type="cellIs" priority="1" operator="equal"><formula>5</formula></cfRule>"#;

        assert_eq!(expected, got);
    }

    #[test]
    fn quoted_string_02() {
        let conditional_format = ConditionalFormatCell::new()
            .set_criteria(ConditionalFormatCellCriteria::EqualTo)
            .set_value("Foo");

        let got = conditional_format.get_rule_string(None, 1);
        let expected = r#"<cfRule type="cellIs" priority="1" operator="equal"><formula>"Foo"</formula></cfRule>"#;

        assert_eq!(expected, got);
    }

    #[test]
    fn quoted_string_03() {
        let conditional_format = ConditionalFormatCell::new()
            .set_criteria(ConditionalFormatCellCriteria::EqualTo)
            .set_value("\"Foo\"");

        let got = conditional_format.get_rule_string(None, 1);
        let expected = r#"<cfRule type="cellIs" priority="1" operator="equal"><formula>"Foo"</formula></cfRule>"#;

        assert_eq!(expected, got);
    }

    #[test]
    fn quoted_string_04() {
        let conditional_format = ConditionalFormatCell::new()
            .set_criteria(ConditionalFormatCellCriteria::EqualTo)
            .set_value("Foo \" Bar");

        let got = conditional_format.get_rule_string(None, 1);
        let expected = r#"<cfRule type="cellIs" priority="1" operator="equal"><formula>"Foo "" Bar"</formula></cfRule>"#;

        assert_eq!(expected, got);
    }

    #[test]
    fn conditional_format_01() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format = ConditionalFormatCell::new()
            .set_criteria(ConditionalFormatCellCriteria::GreaterThan)
            .set_value(5);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:A4"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:1">
                  <c r="A1">
                    <v>10</v>
                  </c>
                </row>
                <row r="2" spans="1:1">
                  <c r="A2">
                    <v>20</v>
                  </c>
                </row>
                <row r="3" spans="1:1">
                  <c r="A3">
                    <v>30</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="A4">
                    <v>40</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A1">
                <cfRule type="cellIs" priority="1" operator="greaterThan">
                  <formula>5</formula>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn conditional_format_02() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        worksheet.write(0, 1, 5)?;

        let conditional_format = ConditionalFormatCell::new()
            .set_criteria(ConditionalFormatCellCriteria::GreaterThan)
            .set_value(Formula::new("$B$1"));

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:B4"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:2">
                  <c r="A1">
                    <v>10</v>
                  </c>
                  <c r="B1">
                    <v>5</v>
                  </c>
                </row>
                <row r="2" spans="1:2">
                  <c r="A2">
                    <v>20</v>
                  </c>
                </row>
                <row r="3" spans="1:2">
                  <c r="A3">
                    <v>30</v>
                  </c>
                </row>
                <row r="4" spans="1:2">
                  <c r="A4">
                    <v>40</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A1">
                <cfRule type="cellIs" priority="1" operator="greaterThan">
                  <formula>$B$1</formula>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn conditional_format_03() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format = ConditionalFormatCell::new()
            .set_criteria(ConditionalFormatCellCriteria::Between)
            .set_minimum(20)
            .set_maximum(30);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatCell::new()
            .set_criteria(ConditionalFormatCellCriteria::NotBetween)
            .set_minimum(20)
            .set_maximum(30)
            .set_multi_range("A1:A4"); // Additional test for multi_range.

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:A4"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:1">
                  <c r="A1">
                    <v>10</v>
                  </c>
                </row>
                <row r="2" spans="1:1">
                  <c r="A2">
                    <v>20</v>
                  </c>
                </row>
                <row r="3" spans="1:1">
                  <c r="A3">
                    <v>30</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="A4">
                    <v>40</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A1:A4">
                <cfRule type="cellIs" priority="1" operator="between">
                  <formula>20</formula>
                  <formula>30</formula>
                </cfRule>
                <cfRule type="cellIs" priority="2" operator="notBetween">
                  <formula>20</formula>
                  <formula>30</formula>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn conditional_format_10() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format = ConditionalFormatCell::new()
            .set_criteria(ConditionalFormatCellCriteria::GreaterThan)
            .set_value(ExcelDateTime::parse_from_str("2024-01-01")?);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:A4"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:1">
                  <c r="A1">
                    <v>10</v>
                  </c>
                </row>
                <row r="2" spans="1:1">
                  <c r="A2">
                    <v>20</v>
                  </c>
                </row>
                <row r="3" spans="1:1">
                  <c r="A3">
                    <v>30</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="A4">
                    <v>40</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A1:A4">
                <cfRule type="cellIs" priority="1" operator="greaterThan">
                  <formula>45292</formula>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn conditional_format_11() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format = ConditionalFormatCell::new()
            .set_minimum(ExcelDateTime::parse_from_str("2024-01-01")?)
            .set_maximum(ExcelDateTime::parse_from_str("2024-01-10")?)
            .set_criteria(ConditionalFormatCellCriteria::Between);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:A4"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:1">
                  <c r="A1">
                    <v>10</v>
                  </c>
                </row>
                <row r="2" spans="1:1">
                  <c r="A2">
                    <v>20</v>
                  </c>
                </row>
                <row r="3" spans="1:1">
                  <c r="A3">
                    <v>30</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="A4">
                    <v>40</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A1:A4">
                <cfRule type="cellIs" priority="1" operator="between">
                  <formula>45292</formula>
                  <formula>45301</formula>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }
}
