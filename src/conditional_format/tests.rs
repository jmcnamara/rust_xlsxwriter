// conditional_format unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod conditional_format_tests {

    use crate::test_functions::xml_to_vec;
    use crate::worksheet::*;
    use crate::ConditionalFormatAverage;
    use crate::ConditionalFormatAverageCriteria;
    use crate::ConditionalFormatCell;
    use crate::ConditionalFormatCellCriteria;
    use crate::ConditionalFormatDate;
    use crate::ConditionalFormatDateCriteria;
    use crate::ConditionalFormatDuplicate;
    use crate::ConditionalFormatText;
    use crate::ConditionalFormatTextCriteria;
    use crate::ConditionalFormatTop;
    use crate::ExcelDateTime;
    use crate::Formula;
    use crate::XlsxError;
    use pretty_assertions::assert_eq;

    #[test]
    fn quoted_string_01() {
        let conditional_format = ConditionalFormatCell::new()
            .set_criteria(ConditionalFormatCellCriteria::EqualTo)
            .set_value(5);

        let got = conditional_format.get_rule_string(None, 1, "");
        let expected =
            r#"<cfRule type="cellIs" priority="1" operator="equal"><formula>5</formula></cfRule>"#;

        assert_eq!(expected, got);
    }

    #[test]
    fn quoted_string_02() {
        let conditional_format = ConditionalFormatCell::new()
            .set_criteria(ConditionalFormatCellCriteria::EqualTo)
            .set_value("Foo");

        let got = conditional_format.get_rule_string(None, 1, "");
        let expected = r#"<cfRule type="cellIs" priority="1" operator="equal"><formula>"Foo"</formula></cfRule>"#;

        assert_eq!(expected, got);
    }

    #[test]
    fn quoted_string_03() {
        let conditional_format = ConditionalFormatCell::new()
            .set_criteria(ConditionalFormatCellCriteria::EqualTo)
            .set_value("\"Foo\"");

        let got = conditional_format.get_rule_string(None, 1, "");
        let expected = r#"<cfRule type="cellIs" priority="1" operator="equal"><formula>"Foo"</formula></cfRule>"#;

        assert_eq!(expected, got);
    }

    #[test]
    fn quoted_string_04() {
        let conditional_format = ConditionalFormatCell::new()
            .set_criteria(ConditionalFormatCellCriteria::EqualTo)
            .set_value("Foo \" Bar");

        let got = conditional_format.get_rule_string(None, 1, "");
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
    fn conditional_format_04() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format = ConditionalFormatDuplicate::new();
        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatDuplicate::new().invert();
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
                <cfRule type="duplicateValues" priority="1"/>
                <cfRule type="uniqueValues" priority="2"/>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn conditional_format_05() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format = ConditionalFormatAverage::new();

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_criteria(ConditionalFormatAverageCriteria::BelowAverage);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_criteria(ConditionalFormatAverageCriteria::EqualOrAboveAverage);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_criteria(ConditionalFormatAverageCriteria::EqualOrBelowAverage);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_criteria(ConditionalFormatAverageCriteria::OneStandardDeviationAbove);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_criteria(ConditionalFormatAverageCriteria::OneStandardDeviationBelow);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_criteria(ConditionalFormatAverageCriteria::TwoStandardDeviationsAbove);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_criteria(ConditionalFormatAverageCriteria::TwoStandardDeviationsBelow);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_criteria(ConditionalFormatAverageCriteria::ThreeStandardDeviationsAbove);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_criteria(ConditionalFormatAverageCriteria::ThreeStandardDeviationsBelow);

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
                <cfRule type="aboveAverage" priority="1"/>
                <cfRule type="aboveAverage" priority="2" aboveAverage="0"/>
                <cfRule type="aboveAverage" priority="3" equalAverage="1"/>
                <cfRule type="aboveAverage" priority="4" aboveAverage="0" equalAverage="1"/>
                <cfRule type="aboveAverage" priority="5" stdDev="1"/>
                <cfRule type="aboveAverage" priority="6" aboveAverage="0" stdDev="1"/>
                <cfRule type="aboveAverage" priority="7" stdDev="2"/>
                <cfRule type="aboveAverage" priority="8" aboveAverage="0" stdDev="2"/>
                <cfRule type="aboveAverage" priority="9" stdDev="3"/>
                <cfRule type="aboveAverage" priority="10" aboveAverage="0" stdDev="3"/>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn conditional_format_06() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format = ConditionalFormatTop::new();
        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatTop::new().invert().set_value(16);
        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatTop::new().set_value(17).set_percent();
        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatTop::new()
            .invert()
            .set_value(18)
            .set_percent();
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
                <cfRule type="top10" priority="1" rank="10"/>
                <cfRule type="top10" priority="2" bottom="1" rank="16"/>
                <cfRule type="top10" priority="3" percent="1" rank="17"/>
                <cfRule type="top10" priority="4" percent="1" bottom="1" rank="18"/>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn conditional_format_07() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format = ConditionalFormatText::new()
            .set_criteria(ConditionalFormatTextCriteria::Contains)
            .set_value("foo");

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_criteria(ConditionalFormatTextCriteria::DoesNotContain)
            .set_value("foo");

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_criteria(ConditionalFormatTextCriteria::BeginsWith)
            .set_value("b");

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_criteria(ConditionalFormatTextCriteria::EndsWith)
            .set_value("b");

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
                <cfRule type="containsText" priority="1" operator="containsText" text="foo">
                  <formula>NOT(ISERROR(SEARCH("foo",A1)))</formula>
                </cfRule>
                <cfRule type="notContainsText" priority="2" operator="notContains" text="foo">
                  <formula>ISERROR(SEARCH("foo",A1))</formula>
                </cfRule>
                <cfRule type="beginsWith" priority="3" operator="beginsWith" text="b">
                  <formula>LEFT(A1,1)="b"</formula>
                </cfRule>
                <cfRule type="endsWith" priority="4" operator="endsWith" text="b">
                  <formula>RIGHT(A1,1)="b"</formula>
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
    fn conditional_format_07b() -> Result<(), XlsxError> {
        // Test different anchor cells.
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format = ConditionalFormatText::new()
            .set_criteria(ConditionalFormatTextCriteria::Contains)
            .set_multi_range("A2")
            .set_value("foo");

        worksheet.add_conditional_format(1, 0, 4, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_criteria(ConditionalFormatTextCriteria::DoesNotContain)
            .set_multi_range("B2:B3")
            .set_value("foo");

        worksheet.add_conditional_format(1, 0, 4, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_criteria(ConditionalFormatTextCriteria::BeginsWith)
            .set_multi_range("C2 C3")
            .set_value("b");

        worksheet.add_conditional_format(1, 0, 4, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_criteria(ConditionalFormatTextCriteria::EndsWith)
            .set_multi_range("D2:D3 D4")
            .set_value("b");

        worksheet.add_conditional_format(1, 0, 4, 0, &conditional_format)?;

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
              <conditionalFormatting sqref="A2">
              <cfRule type="containsText" priority="1" operator="containsText" text="foo">
              <formula>NOT(ISERROR(SEARCH("foo",A2)))</formula>
              </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="B2:B3">
              <cfRule type="notContainsText" priority="2" operator="notContains" text="foo">
              <formula>ISERROR(SEARCH("foo",B2))</formula>
              </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="C2 C3">
              <cfRule type="beginsWith" priority="3" operator="beginsWith" text="b">
              <formula>LEFT(C2,1)="b"</formula>
              </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="D2:D3 D4">
                <cfRule type="endsWith" priority="4" operator="endsWith" text="b">
                  <formula>RIGHT(D2,1)="b"</formula>
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
    fn conditional_format_08() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format =
            ConditionalFormatDate::new().set_criteria(ConditionalFormatDateCriteria::Yesterday);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_criteria(ConditionalFormatDateCriteria::Today);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_criteria(ConditionalFormatDateCriteria::Tomorrow);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_criteria(ConditionalFormatDateCriteria::Last7Days);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_criteria(ConditionalFormatDateCriteria::LastWeek);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_criteria(ConditionalFormatDateCriteria::ThisWeek);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_criteria(ConditionalFormatDateCriteria::NextWeek);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_criteria(ConditionalFormatDateCriteria::LastMonth);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_criteria(ConditionalFormatDateCriteria::ThisMonth);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_criteria(ConditionalFormatDateCriteria::NextMonth);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

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
              <conditionalFormatting sqref="C2:C5">
                <cfRule type="timePeriod" priority="1" timePeriod="yesterday">
                  <formula>FLOOR(C2,1)=TODAY()-1</formula>
                </cfRule>
                <cfRule type="timePeriod" priority="2" timePeriod="today">
                  <formula>FLOOR(C2,1)=TODAY()</formula>
                </cfRule>
                <cfRule type="timePeriod" priority="3" timePeriod="tomorrow">
                  <formula>FLOOR(C2,1)=TODAY()+1</formula>
                </cfRule>
                <cfRule type="timePeriod" priority="4" timePeriod="last7Days">
                  <formula>AND(TODAY()-FLOOR(C2,1)&lt;=6,FLOOR(C2,1)&lt;=TODAY())</formula>
                </cfRule>
                <cfRule type="timePeriod" priority="5" timePeriod="lastWeek">
                  <formula>AND(TODAY()-ROUNDDOWN(C2,0)&gt;=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(C2,0)&lt;(WEEKDAY(TODAY())+7))</formula>
                </cfRule>
                <cfRule type="timePeriod" priority="6" timePeriod="thisWeek">
                  <formula>AND(TODAY()-ROUNDDOWN(C2,0)&lt;=WEEKDAY(TODAY())-1,ROUNDDOWN(C2,0)-TODAY()&lt;=7-WEEKDAY(TODAY()))</formula>
                </cfRule>
                <cfRule type="timePeriod" priority="7" timePeriod="nextWeek">
                  <formula>AND(ROUNDDOWN(C2,0)-TODAY()&gt;(7-WEEKDAY(TODAY())),ROUNDDOWN(C2,0)-TODAY()&lt;(15-WEEKDAY(TODAY())))</formula>
                </cfRule>
                <cfRule type="timePeriod" priority="8" timePeriod="lastMonth">
                  <formula>AND(MONTH(C2)=MONTH(TODAY())-1,OR(YEAR(C2)=YEAR(TODAY()),AND(MONTH(C2)=1,YEAR(C2)=YEAR(TODAY())-1)))</formula>
                </cfRule>
                <cfRule type="timePeriod" priority="9" timePeriod="thisMonth">
                  <formula>AND(MONTH(C2)=MONTH(TODAY()),YEAR(C2)=YEAR(TODAY()))</formula>
                </cfRule>
                <cfRule type="timePeriod" priority="10" timePeriod="nextMonth">
                  <formula>AND(MONTH(C2)=MONTH(TODAY())+1,OR(YEAR(C2)=YEAR(TODAY()),AND(MONTH(C2)=12,YEAR(C2)=YEAR(TODAY())+1)))</formula>
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

    #[test]
    fn validation_checks() {
        // Check validations for various conditional formats.

        // Cell format must have a non-None type.
        let conditional_format = ConditionalFormatCell::new();
        let result = conditional_format.validate();
        assert!(matches!(result, Err(XlsxError::ConditionalFormatError(_))));

        // Cell format must have a value.
        let conditional_format =
            ConditionalFormatCell::new().set_criteria(ConditionalFormatCellCriteria::EqualTo);
        let result = conditional_format.validate();
        assert!(matches!(result, Err(XlsxError::ConditionalFormatError(_))));

        // Cell between format must have a min value.
        let conditional_format =
            ConditionalFormatCell::new().set_criteria(ConditionalFormatCellCriteria::Between);
        let result = conditional_format.validate();
        assert!(matches!(result, Err(XlsxError::ConditionalFormatError(_))));

        // Cell between format must have a max value.
        let conditional_format = ConditionalFormatCell::new()
            .set_criteria(ConditionalFormatCellCriteria::Between)
            .set_minimum(1);
        let result = conditional_format.validate();
        assert!(matches!(result, Err(XlsxError::ConditionalFormatError(_))));

        // Top value must be in the Excel range 1..1000.
        let conditional_format = ConditionalFormatTop::new().set_value(0);
        let result = conditional_format.validate();
        assert!(matches!(result, Err(XlsxError::ConditionalFormatError(_))));

        // Top value must be in the Excel range 1..1000.
        let conditional_format = ConditionalFormatTop::new().set_value(1001);
        let result = conditional_format.validate();
        assert!(matches!(result, Err(XlsxError::ConditionalFormatError(_))));
    }

    #[test]
    fn multi_range_replacing() {
        // Check escaping of the multi-range string.

        let conditional_format = ConditionalFormatCell::new().set_multi_range("A1");
        let multi_range = conditional_format.multi_range();
        assert_eq!("A1", multi_range);

        let conditional_format = ConditionalFormatCell::new().set_multi_range("$A$1");
        let multi_range = conditional_format.multi_range();
        assert_eq!("A1", multi_range);

        let conditional_format = ConditionalFormatCell::new()
            .set_multi_range("$B$3:$D$6,$I$3:$K$6,$B$9:$D$12,$I$9:$K$12");
        let multi_range = conditional_format.multi_range();
        assert_eq!("B3:D6 I3:K6 B9:D12 I9:K12", multi_range);
    }
}
