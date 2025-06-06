// conditional_format unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod conditional_format_tests {

    use crate::test_functions::xml_to_vec;
    use crate::worksheet::*;
    use crate::xmlwriter;
    use crate::ConditionalFormat2ColorScale;
    use crate::ConditionalFormat3ColorScale;
    use crate::ConditionalFormatAverage;
    use crate::ConditionalFormatAverageRule;
    use crate::ConditionalFormatBlank;
    use crate::ConditionalFormatCell;
    use crate::ConditionalFormatCellRule;
    use crate::ConditionalFormatCustomIcon;
    use crate::ConditionalFormatDataBar;
    use crate::ConditionalFormatDataBarAxisPosition;
    use crate::ConditionalFormatDataBarDirection;
    use crate::ConditionalFormatDate;
    use crate::ConditionalFormatDateRule;
    use crate::ConditionalFormatDuplicate;
    use crate::ConditionalFormatError;
    use crate::ConditionalFormatFormula;
    use crate::ConditionalFormatIconSet;
    use crate::ConditionalFormatIconType;
    use crate::ConditionalFormatText;
    use crate::ConditionalFormatTextRule;
    use crate::ConditionalFormatTop;
    use crate::ConditionalFormatTopRule;
    use crate::ConditionalFormatType;
    use crate::ExcelDateTime;
    use crate::Formula;
    use crate::XlsxError;
    use pretty_assertions::assert_eq;

    #[test]
    fn quoted_string_01() {
        let conditional_format =
            ConditionalFormatCell::new().set_rule(ConditionalFormatCellRule::EqualTo(5));

        let got = conditional_format.rule(None, 1, "", "");
        let expected =
            r#"<cfRule type="cellIs" priority="1" operator="equal"><formula>5</formula></cfRule>"#;

        assert_eq!(expected, got);
    }

    #[test]
    fn quoted_string_02() {
        let conditional_format =
            ConditionalFormatCell::new().set_rule(ConditionalFormatCellRule::EqualTo("Foo"));

        let got = conditional_format.rule(None, 1, "", "");
        let expected = r#"<cfRule type="cellIs" priority="1" operator="equal"><formula>"Foo"</formula></cfRule>"#;

        assert_eq!(expected, got);
    }

    #[test]
    fn quoted_string_03() {
        let conditional_format =
            ConditionalFormatCell::new().set_rule(ConditionalFormatCellRule::EqualTo("\"Foo\""));

        let got = conditional_format.rule(None, 1, "", "");
        let expected = r#"<cfRule type="cellIs" priority="1" operator="equal"><formula>"Foo"</formula></cfRule>"#;

        assert_eq!(expected, got);
    }

    #[test]
    fn quoted_string_04() {
        let conditional_format =
            ConditionalFormatCell::new().set_rule(ConditionalFormatCellRule::EqualTo("Foo \" Bar"));

        let got = conditional_format.rule(None, 1, "", "");
        let expected = r#"<cfRule type="cellIs" priority="1" operator="equal"><formula>"Foo "" Bar"</formula></cfRule>"#;

        assert_eq!(expected, got);
    }

    #[test]
    fn validation_checks() {
        // Check validations for various conditional formats.

        // Cell format must have a rule.
        let conditional_format = ConditionalFormatCell::new();
        let result = conditional_format.validate();
        assert!(matches!(result, Err(XlsxError::ConditionalFormatError(_))));

        // Top value must be in the Excel range 1..1000.
        let conditional_format =
            ConditionalFormatTop::new().set_rule(ConditionalFormatTopRule::Top(0));
        let result = conditional_format.validate();
        assert!(matches!(result, Err(XlsxError::ConditionalFormatError(_))));

        // Top value must be in the Excel range 1..1000.
        let conditional_format =
            ConditionalFormatTop::new().set_rule(ConditionalFormatTopRule::Top(1001));
        let result = conditional_format.validate();
        assert!(matches!(result, Err(XlsxError::ConditionalFormatError(_))));

        // Check validation of the number of user supplied rules. 3 icon style.
        let icons = [
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 1),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 20),
        ];
        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeArrows)
            .set_icons(&icons);
        let result = conditional_format.validate();
        assert!(matches!(result, Err(XlsxError::ConditionalFormatError(_))));

        // Check validation of the number of user supplied rules. 4 icon style.
        let icons = [
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 1),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 2),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 3),
        ];
        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::FourArrows)
            .set_icons(&icons);
        let result = conditional_format.validate();
        assert!(matches!(result, Err(XlsxError::ConditionalFormatError(_))));

        // Check validation of the number of user supplied rules. 5 icon style.
        let icons = [
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 1),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 2),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 3),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 4),
        ];
        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::FiveArrows)
            .set_icons(&icons);
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

    #[test]
    fn conditional_format_01() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format =
            ConditionalFormatCell::new().set_rule(ConditionalFormatCellRule::GreaterThan(5));

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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
            .set_rule(ConditionalFormatCellRule::GreaterThan(Formula::new("$B$1")));

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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

        let conditional_format =
            ConditionalFormatCell::new().set_rule(ConditionalFormatCellRule::Between(20, 30));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatCell::new()
            .set_rule(ConditionalFormatCellRule::NotBetween(20, 30))
            .set_multi_range("A1:A4"); // Additional test for multi_range.

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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

        let conditional_format =
            ConditionalFormatAverage::new().set_rule(ConditionalFormatAverageRule::BelowAverage);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_rule(ConditionalFormatAverageRule::EqualOrAboveAverage);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_rule(ConditionalFormatAverageRule::EqualOrBelowAverage);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_rule(ConditionalFormatAverageRule::OneStandardDeviationAbove);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_rule(ConditionalFormatAverageRule::OneStandardDeviationBelow);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_rule(ConditionalFormatAverageRule::TwoStandardDeviationsAbove);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_rule(ConditionalFormatAverageRule::TwoStandardDeviationsBelow);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_rule(ConditionalFormatAverageRule::ThreeStandardDeviationsAbove);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatAverage::new()
            .set_rule(ConditionalFormatAverageRule::ThreeStandardDeviationsBelow);

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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

        let conditional_format =
            ConditionalFormatTop::new().set_rule(ConditionalFormatTopRule::Bottom(16));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format =
            ConditionalFormatTop::new().set_rule(ConditionalFormatTopRule::TopPercent(17));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format =
            ConditionalFormatTop::new().set_rule(ConditionalFormatTopRule::BottomPercent(18));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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
            .set_rule(ConditionalFormatTextRule::Contains("foo".to_string()));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_rule(ConditionalFormatTextRule::DoesNotContain("foo".to_string()));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_rule(ConditionalFormatTextRule::BeginsWith("b".to_string()));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_rule(ConditionalFormatTextRule::EndsWith("b".to_string()));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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
            .set_rule(ConditionalFormatTextRule::Contains("foo".to_string()))
            .set_multi_range("A2");

        worksheet.add_conditional_format(1, 0, 4, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_rule(ConditionalFormatTextRule::DoesNotContain("foo".to_string()))
            .set_multi_range("B2:B3");

        worksheet.add_conditional_format(1, 0, 4, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_rule(ConditionalFormatTextRule::BeginsWith("b".to_string()))
            .set_multi_range("C2 C3");

        worksheet.add_conditional_format(1, 0, 4, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_rule(ConditionalFormatTextRule::EndsWith("b".to_string()))
            .set_multi_range("D2:D3 D4");

        worksheet.add_conditional_format(1, 0, 4, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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
            ConditionalFormatDate::new().set_rule(ConditionalFormatDateRule::Yesterday);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_rule(ConditionalFormatDateRule::Today);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_rule(ConditionalFormatDateRule::Tomorrow);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_rule(ConditionalFormatDateRule::Last7Days);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_rule(ConditionalFormatDateRule::LastWeek);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_rule(ConditionalFormatDateRule::ThisWeek);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_rule(ConditionalFormatDateRule::NextWeek);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_rule(ConditionalFormatDateRule::LastMonth);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_rule(ConditionalFormatDateRule::ThisMonth);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        let conditional_format =
            ConditionalFormatDate::new().set_rule(ConditionalFormatDateRule::NextMonth);
        worksheet.add_conditional_format(1, 2, 4, 2, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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
    fn conditional_format_09() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format = ConditionalFormatBlank::new();
        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatBlank::new().invert();
        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatError::new();
        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatError::new().invert();
        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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
                <cfRule type="containsBlanks" priority="1">
                  <formula>LEN(TRIM(A1))=0</formula>
                </cfRule>
                <cfRule type="notContainsBlanks" priority="2">
                  <formula>LEN(TRIM(A1))&gt;0</formula>
                </cfRule>
                <cfRule type="containsErrors" priority="3">
                  <formula>ISERROR(A1)</formula>
                </cfRule>
                <cfRule type="notContainsErrors" priority="4">
                  <formula>NOT(ISERROR(A1))</formula>
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

        let conditional_format = ConditionalFormatCell::new().set_rule(
            ConditionalFormatCellRule::GreaterThan(ExcelDateTime::parse_from_str("2024-01-01")?),
        );

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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

        let conditional_format =
            ConditionalFormatCell::new().set_rule(ConditionalFormatCellRule::Between(
                ExcelDateTime::parse_from_str("2024-01-01")?,
                ExcelDateTime::parse_from_str("2024-01-10")?,
            ));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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

    #[cfg(feature = "chrono")]
    #[test]
    fn conditional_format_11_chrono() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format =
            ConditionalFormatCell::new().set_rule(ConditionalFormatCellRule::Between(
                chrono::NaiveDate::from_ymd_opt(2024, 1, 1).unwrap(),
                chrono::NaiveDate::from_ymd_opt(2024, 1, 10).unwrap(),
            ));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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

    #[cfg(feature = "jiff")]
    #[test]
    fn conditional_format_11_jiff() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format =
            ConditionalFormatCell::new().set_rule(ConditionalFormatCellRule::Between(
                jiff::civil::date(2024, 1, 1),
                jiff::civil::date(2024, 1, 10),
            ));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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
    fn conditional_format_12() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 1)?;
        worksheet.write(1, 0, 2)?;
        worksheet.write(2, 0, 3)?;
        worksheet.write(3, 0, 4)?;
        worksheet.write(4, 0, 5)?;
        worksheet.write(5, 0, 6)?;
        worksheet.write(6, 0, 7)?;
        worksheet.write(7, 0, 8)?;
        worksheet.write(8, 0, 9)?;
        worksheet.write(9, 0, 10)?;
        worksheet.write(10, 0, 11)?;
        worksheet.write(11, 0, 12)?;

        let conditional_format = ConditionalFormat2ColorScale::new();

        worksheet.add_conditional_format(0, 0, 11, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:A12"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:1">
                  <c r="A1">
                    <v>1</v>
                  </c>
                </row>
                <row r="2" spans="1:1">
                  <c r="A2">
                    <v>2</v>
                  </c>
                </row>
                <row r="3" spans="1:1">
                  <c r="A3">
                    <v>3</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="A4">
                    <v>4</v>
                  </c>
                </row>
                <row r="5" spans="1:1">
                  <c r="A5">
                    <v>5</v>
                  </c>
                </row>
                <row r="6" spans="1:1">
                  <c r="A6">
                    <v>6</v>
                  </c>
                </row>
                <row r="7" spans="1:1">
                  <c r="A7">
                    <v>7</v>
                  </c>
                </row>
                <row r="8" spans="1:1">
                  <c r="A8">
                    <v>8</v>
                  </c>
                </row>
                <row r="9" spans="1:1">
                  <c r="A9">
                    <v>9</v>
                  </c>
                </row>
                <row r="10" spans="1:1">
                  <c r="A10">
                    <v>10</v>
                  </c>
                </row>
                <row r="11" spans="1:1">
                  <c r="A11">
                    <v>11</v>
                  </c>
                </row>
                <row r="12" spans="1:1">
                  <c r="A12">
                    <v>12</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A1:A12">
                <cfRule type="colorScale" priority="1">
                  <colorScale>
                    <cfvo type="min" val="0"/>
                    <cfvo type="max" val="0"/>
                    <color rgb="FFFFEF9C"/>
                    <color rgb="FF63BE7B"/>
                  </colorScale>
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
    fn conditional_format_12b() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormat2ColorScale::new()
            // String should be ignored.
            .set_minimum(ConditionalFormatType::Number, "Foo")
            .set_maximum(ConditionalFormatType::Number, "Foo")
            // High/low should be ignored.
            .set_minimum(ConditionalFormatType::Highest, 0)
            .set_maximum(ConditionalFormatType::Lowest, 0)
            // > 100 should be ignored.
            .set_minimum(ConditionalFormatType::Percent, 101)
            .set_maximum(ConditionalFormatType::Percent, 101)
            // < 0 should be ignored.
            .set_minimum(ConditionalFormatType::Percentile, -1)
            .set_maximum(ConditionalFormatType::Percentile, -1)
            .set_minimum_color("FF0000")
            .set_maximum_color("FFFF00");
        worksheet.add_conditional_format(0, 0, 9, 0, &conditional_format)?;

        let conditional_format = ConditionalFormat2ColorScale::new()
            .set_minimum(ConditionalFormatType::Number, 2.5)
            .set_maximum(ConditionalFormatType::Percent, 90);
        worksheet.add_conditional_format(0, 2, 9, 2, &conditional_format)?;

        let conditional_format = ConditionalFormat2ColorScale::new()
            .set_minimum(ConditionalFormatType::Formula, Formula::new("=$M$20"))
            .set_maximum(ConditionalFormatType::Percentile, 90);
        worksheet.add_conditional_format(0, 4, 9, 4, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData/>
              <conditionalFormatting sqref="A1:A10">
                <cfRule type="colorScale" priority="1">
                  <colorScale>
                    <cfvo type="min" val="0"/>
                    <cfvo type="max" val="0"/>
                    <color rgb="FFFF0000"/>
                    <color rgb="FFFFFF00"/>
                  </colorScale>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="C1:C10">
                <cfRule type="colorScale" priority="2">
                  <colorScale>
                    <cfvo type="num" val="2.5"/>
                    <cfvo type="percent" val="90"/>
                    <color rgb="FFFFEF9C"/>
                    <color rgb="FF63BE7B"/>
                  </colorScale>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="E1:E10">
                <cfRule type="colorScale" priority="3">
                  <colorScale>
                    <cfvo type="formula" val="$M$20"/>
                    <cfvo type="percentile" val="90"/>
                    <color rgb="FFFFEF9C"/>
                    <color rgb="FF63BE7B"/>
                  </colorScale>
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
    fn conditional_format_13() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 1)?;
        worksheet.write(1, 0, 2)?;
        worksheet.write(2, 0, 3)?;
        worksheet.write(3, 0, 4)?;
        worksheet.write(4, 0, 5)?;
        worksheet.write(5, 0, 6)?;
        worksheet.write(6, 0, 7)?;
        worksheet.write(7, 0, 8)?;
        worksheet.write(8, 0, 9)?;
        worksheet.write(9, 0, 10)?;
        worksheet.write(10, 0, 11)?;
        worksheet.write(11, 0, 12)?;

        let conditional_format = ConditionalFormat3ColorScale::new()
            .set_minimum_color("F8696B")
            .set_midpoint_color("FFEB84")
            .set_maximum_color("63BE7B");

        worksheet.add_conditional_format(0, 0, 11, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:A12"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:1">
                  <c r="A1">
                    <v>1</v>
                  </c>
                </row>
                <row r="2" spans="1:1">
                  <c r="A2">
                    <v>2</v>
                  </c>
                </row>
                <row r="3" spans="1:1">
                  <c r="A3">
                    <v>3</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="A4">
                    <v>4</v>
                  </c>
                </row>
                <row r="5" spans="1:1">
                  <c r="A5">
                    <v>5</v>
                  </c>
                </row>
                <row r="6" spans="1:1">
                  <c r="A6">
                    <v>6</v>
                  </c>
                </row>
                <row r="7" spans="1:1">
                  <c r="A7">
                    <v>7</v>
                  </c>
                </row>
                <row r="8" spans="1:1">
                  <c r="A8">
                    <v>8</v>
                  </c>
                </row>
                <row r="9" spans="1:1">
                  <c r="A9">
                    <v>9</v>
                  </c>
                </row>
                <row r="10" spans="1:1">
                  <c r="A10">
                    <v>10</v>
                  </c>
                </row>
                <row r="11" spans="1:1">
                  <c r="A11">
                    <v>11</v>
                  </c>
                </row>
                <row r="12" spans="1:1">
                  <c r="A12">
                    <v>12</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A1:A12">
                <cfRule type="colorScale" priority="1">
                  <colorScale>
                    <cfvo type="min" val="0"/>
                    <cfvo type="percentile" val="50"/>
                    <cfvo type="max" val="0"/>
                    <color rgb="FFF8696B"/>
                    <color rgb="FFFFEB84"/>
                    <color rgb="FF63BE7B"/>
                  </colorScale>
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
    fn conditional_format_14() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 1)?;
        worksheet.write(1, 0, 2)?;
        worksheet.write(2, 0, 3)?;
        worksheet.write(3, 0, 4)?;
        worksheet.write(4, 0, 5)?;
        worksheet.write(5, 0, 6)?;
        worksheet.write(6, 0, 7)?;
        worksheet.write(7, 0, 8)?;
        worksheet.write(8, 0, 9)?;
        worksheet.write(9, 0, 10)?;
        worksheet.write(10, 0, 11)?;
        worksheet.write(11, 0, 12)?;

        let conditional_format = ConditionalFormatDataBar::new().use_classic_style();

        worksheet.add_conditional_format(0, 0, 11, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:A12"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:1">
                  <c r="A1">
                    <v>1</v>
                  </c>
                </row>
                <row r="2" spans="1:1">
                  <c r="A2">
                    <v>2</v>
                  </c>
                </row>
                <row r="3" spans="1:1">
                  <c r="A3">
                    <v>3</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="A4">
                    <v>4</v>
                  </c>
                </row>
                <row r="5" spans="1:1">
                  <c r="A5">
                    <v>5</v>
                  </c>
                </row>
                <row r="6" spans="1:1">
                  <c r="A6">
                    <v>6</v>
                  </c>
                </row>
                <row r="7" spans="1:1">
                  <c r="A7">
                    <v>7</v>
                  </c>
                </row>
                <row r="8" spans="1:1">
                  <c r="A8">
                    <v>8</v>
                  </c>
                </row>
                <row r="9" spans="1:1">
                  <c r="A9">
                    <v>9</v>
                  </c>
                </row>
                <row r="10" spans="1:1">
                  <c r="A10">
                    <v>10</v>
                  </c>
                </row>
                <row r="11" spans="1:1">
                  <c r="A11">
                    <v>11</v>
                  </c>
                </row>
                <row r="12" spans="1:1">
                  <c r="A12">
                    <v>12</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A1:A12">
                <cfRule type="dataBar" priority="1">
                  <dataBar>
                    <cfvo type="min" val="0"/>
                    <cfvo type="max" val="0"/>
                    <color rgb="FF638EC6"/>
                  </dataBar>
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
    fn conditional_format_15() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format = ConditionalFormatFormula::new().set_rule("=$A$1>5");

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatFormula::new().set_rule("=$A$2<80");

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatFormula::new().set_rule("\"1+2\"");

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatFormula::new().set_rule("=$A$3>$A$4");

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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
                <cfRule type="expression" priority="1">
                  <formula>$A$1&gt;5</formula>
                </cfRule>
                <cfRule type="expression" priority="2">
                  <formula>$A$2&lt;80</formula>
                </cfRule>
                <cfRule type="expression" priority="3">
                  <formula>"1+2"</formula>
                </cfRule>
                <cfRule type="expression" priority="4">
                  <formula>$A$3&gt;$A$4</formula>
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
    fn conditional_format_16() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 1)?;
        worksheet.write(1, 0, 2)?;
        worksheet.write(2, 0, 3)?;
        worksheet.write(3, 0, 4)?;
        worksheet.write(4, 0, 5)?;
        worksheet.write(5, 0, 6)?;
        worksheet.write(6, 0, 7)?;
        worksheet.write(7, 0, 8)?;
        worksheet.write(8, 0, 9)?;
        worksheet.write(9, 0, 10)?;
        worksheet.write(10, 0, 11)?;
        worksheet.write(11, 0, 12)?;

        let conditional_format = ConditionalFormat3ColorScale::new()
            .set_minimum_color("C5D9F1")
            .set_midpoint_color("8DB4E3")
            .set_maximum_color("538ED5");

        worksheet.add_conditional_format(0, 0, 11, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:A12"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:1">
                  <c r="A1">
                    <v>1</v>
                  </c>
                </row>
                <row r="2" spans="1:1">
                  <c r="A2">
                    <v>2</v>
                  </c>
                </row>
                <row r="3" spans="1:1">
                  <c r="A3">
                    <v>3</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="A4">
                    <v>4</v>
                  </c>
                </row>
                <row r="5" spans="1:1">
                  <c r="A5">
                    <v>5</v>
                  </c>
                </row>
                <row r="6" spans="1:1">
                  <c r="A6">
                    <v>6</v>
                  </c>
                </row>
                <row r="7" spans="1:1">
                  <c r="A7">
                    <v>7</v>
                  </c>
                </row>
                <row r="8" spans="1:1">
                  <c r="A8">
                    <v>8</v>
                  </c>
                </row>
                <row r="9" spans="1:1">
                  <c r="A9">
                    <v>9</v>
                  </c>
                </row>
                <row r="10" spans="1:1">
                  <c r="A10">
                    <v>10</v>
                  </c>
                </row>
                <row r="11" spans="1:1">
                  <c r="A11">
                    <v>11</v>
                  </c>
                </row>
                <row r="12" spans="1:1">
                  <c r="A12">
                    <v>12</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A1:A12">
                <cfRule type="colorScale" priority="1">
                  <colorScale>
                    <cfvo type="min" val="0"/>
                    <cfvo type="percentile" val="50"/>
                    <cfvo type="max" val="0"/>
                    <color rgb="FFC5D9F1"/>
                    <color rgb="FF8DB4E3"/>
                    <color rgb="FF538ED5"/>
                  </colorScale>
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
    fn conditional_format_17() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 1)?;
        worksheet.write(1, 0, 2)?;
        worksheet.write(2, 0, 3)?;
        worksheet.write(3, 0, 4)?;
        worksheet.write(4, 0, 5)?;
        worksheet.write(5, 0, 6)?;
        worksheet.write(6, 0, 7)?;
        worksheet.write(7, 0, 8)?;
        worksheet.write(8, 0, 9)?;
        worksheet.write(9, 0, 10)?;
        worksheet.write(10, 0, 11)?;
        worksheet.write(11, 0, 12)?;

        let conditional_format = ConditionalFormat3ColorScale::new()
            .set_minimum(ConditionalFormatType::Number, Formula::new("$A$10"))
            .set_midpoint(ConditionalFormatType::Percent, 52)
            .set_maximum(ConditionalFormatType::Percentile, 99);

        worksheet.add_conditional_format(0, 0, 11, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:A12"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:1">
                  <c r="A1">
                    <v>1</v>
                  </c>
                </row>
                <row r="2" spans="1:1">
                  <c r="A2">
                    <v>2</v>
                  </c>
                </row>
                <row r="3" spans="1:1">
                  <c r="A3">
                    <v>3</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="A4">
                    <v>4</v>
                  </c>
                </row>
                <row r="5" spans="1:1">
                  <c r="A5">
                    <v>5</v>
                  </c>
                </row>
                <row r="6" spans="1:1">
                  <c r="A6">
                    <v>6</v>
                  </c>
                </row>
                <row r="7" spans="1:1">
                  <c r="A7">
                    <v>7</v>
                  </c>
                </row>
                <row r="8" spans="1:1">
                  <c r="A8">
                    <v>8</v>
                  </c>
                </row>
                <row r="9" spans="1:1">
                  <c r="A9">
                    <v>9</v>
                  </c>
                </row>
                <row r="10" spans="1:1">
                  <c r="A10">
                    <v>10</v>
                  </c>
                </row>
                <row r="11" spans="1:1">
                  <c r="A11">
                    <v>11</v>
                  </c>
                </row>
                <row r="12" spans="1:1">
                  <c r="A12">
                    <v>12</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A1:A12">
                <cfRule type="colorScale" priority="1">
                  <colorScale>
                    <cfvo type="num" val="$A$10"/>
                    <cfvo type="percent" val="52"/>
                    <cfvo type="percentile" val="99"/>
                    <color rgb="FFF8696B"/>
                    <color rgb="FFFFEB84"/>
                    <color rgb="FF63BE7B"/>
                  </colorScale>
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
    fn conditional_format_18() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 1)?;
        worksheet.write(1, 0, 2)?;
        worksheet.write(2, 0, 3)?;
        worksheet.write(3, 0, 4)?;
        worksheet.write(4, 0, 5)?;
        worksheet.write(5, 0, 6)?;
        worksheet.write(6, 0, 7)?;
        worksheet.write(7, 0, 8)?;
        worksheet.write(8, 0, 9)?;
        worksheet.write(9, 0, 10)?;
        worksheet.write(10, 0, 11)?;
        worksheet.write(11, 0, 12)?;

        let conditional_format =
            ConditionalFormat3ColorScale::new().set_multi_range("$A$3:$A$4,A1,A6:$A$8,$A10,A$12");

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:A12"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:1">
                  <c r="A1">
                    <v>1</v>
                  </c>
                </row>
                <row r="2" spans="1:1">
                  <c r="A2">
                    <v>2</v>
                  </c>
                </row>
                <row r="3" spans="1:1">
                  <c r="A3">
                    <v>3</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="A4">
                    <v>4</v>
                  </c>
                </row>
                <row r="5" spans="1:1">
                  <c r="A5">
                    <v>5</v>
                  </c>
                </row>
                <row r="6" spans="1:1">
                  <c r="A6">
                    <v>6</v>
                  </c>
                </row>
                <row r="7" spans="1:1">
                  <c r="A7">
                    <v>7</v>
                  </c>
                </row>
                <row r="8" spans="1:1">
                  <c r="A8">
                    <v>8</v>
                  </c>
                </row>
                <row r="9" spans="1:1">
                  <c r="A9">
                    <v>9</v>
                  </c>
                </row>
                <row r="10" spans="1:1">
                  <c r="A10">
                    <v>10</v>
                  </c>
                </row>
                <row r="11" spans="1:1">
                  <c r="A11">
                    <v>11</v>
                  </c>
                </row>
                <row r="12" spans="1:1">
                  <c r="A12">
                    <v>12</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A3:A4 A1 A6:A8 A10 A12">
                <cfRule type="colorScale" priority="1">
                  <colorScale>
                    <cfvo type="min" val="0"/>
                    <cfvo type="percentile" val="50"/>
                    <cfvo type="max" val="0"/>
                    <color rgb="FFF8696B"/>
                    <color rgb="FFFFEB84"/>
                    <color rgb="FF63BE7B"/>
                  </colorScale>
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
    fn conditional_format_19() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 1)?;
        worksheet.write(1, 0, 2)?;
        worksheet.write(2, 0, 3)?;
        worksheet.write(3, 0, 4)?;
        worksheet.write(4, 0, 5)?;
        worksheet.write(5, 0, 6)?;
        worksheet.write(6, 0, 7)?;
        worksheet.write(7, 0, 8)?;
        worksheet.write(8, 0, 9)?;
        worksheet.write(9, 0, 10)?;
        worksheet.write(10, 0, 11)?;
        worksheet.write(11, 0, 12)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_minimum(ConditionalFormatType::Number, 5)
            .set_maximum(ConditionalFormatType::Percent, 90)
            .set_fill_color("8DB4E3")
            .use_classic_style();

        worksheet.add_conditional_format(0, 0, 11, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:A12"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:1">
                  <c r="A1">
                    <v>1</v>
                  </c>
                </row>
                <row r="2" spans="1:1">
                  <c r="A2">
                    <v>2</v>
                  </c>
                </row>
                <row r="3" spans="1:1">
                  <c r="A3">
                    <v>3</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="A4">
                    <v>4</v>
                  </c>
                </row>
                <row r="5" spans="1:1">
                  <c r="A5">
                    <v>5</v>
                  </c>
                </row>
                <row r="6" spans="1:1">
                  <c r="A6">
                    <v>6</v>
                  </c>
                </row>
                <row r="7" spans="1:1">
                  <c r="A7">
                    <v>7</v>
                  </c>
                </row>
                <row r="8" spans="1:1">
                  <c r="A8">
                    <v>8</v>
                  </c>
                </row>
                <row r="9" spans="1:1">
                  <c r="A9">
                    <v>9</v>
                  </c>
                </row>
                <row r="10" spans="1:1">
                  <c r="A10">
                    <v>10</v>
                  </c>
                </row>
                <row r="11" spans="1:1">
                  <c r="A11">
                    <v>11</v>
                  </c>
                </row>
                <row r="12" spans="1:1">
                  <c r="A12">
                    <v>12</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A1:A12">
                <cfRule type="dataBar" priority="1">
                  <dataBar>
                    <cfvo type="num" val="5"/>
                    <cfvo type="percent" val="90"/>
                    <color rgb="FF8DB4E3"/>
                  </dataBar>
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
    fn conditional_format_20() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format = ConditionalFormatText::new()
            .set_rule(ConditionalFormatTextRule::BeginsWith("b".to_string()));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_rule(ConditionalFormatTextRule::BeginsWith("bc".to_string()));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_rule(ConditionalFormatTextRule::EndsWith("z".to_string()));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatText::new()
            .set_rule(ConditionalFormatTextRule::EndsWith("yz".to_string()));

        worksheet.add_conditional_format(0, 0, 3, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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
                <cfRule type="beginsWith" priority="1" operator="beginsWith" text="b">
                  <formula>LEFT(A1,1)="b"</formula>
                </cfRule>
                <cfRule type="beginsWith" priority="2" operator="beginsWith" text="bc">
                  <formula>LEFT(A1,2)="bc"</formula>
                </cfRule>
                <cfRule type="endsWith" priority="3" operator="endsWith" text="z">
                  <formula>RIGHT(A1,1)="z"</formula>
                </cfRule>
                <cfRule type="endsWith" priority="4" operator="endsWith" text="yz">
                  <formula>RIGHT(A1,2)="yz"</formula>
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
    fn conditional_format_21() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 10)?;
        worksheet.write(1, 0, 20)?;
        worksheet.write(2, 0, 30)?;
        worksheet.write(3, 0, 40)?;

        let conditional_format = ConditionalFormatCell::new()
            .set_rule(ConditionalFormatCellRule::GreaterThan(5))
            .set_stop_if_true(true);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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
                <cfRule type="cellIs" priority="1" stopIfTrue="1" operator="greaterThan">
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
    fn conditional_format_22() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 1)?;
        worksheet.write(1, 0, 2)?;
        worksheet.write(2, 0, 3)?;
        worksheet.write(3, 0, 4)?;
        worksheet.write(4, 0, 5)?;
        worksheet.write(5, 0, 6)?;
        worksheet.write(6, 0, 7)?;
        worksheet.write(7, 0, 8)?;
        worksheet.write(8, 0, 9)?;

        let conditional_format =
            ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::ThreeArrows);
        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format =
            ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::ThreeFlags);
        worksheet.add_conditional_format(1, 0, 1, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeTrafficLightsWithRim);
        worksheet.add_conditional_format(2, 0, 2, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeSymbolsCircled);
        worksheet.add_conditional_format(3, 0, 3, 0, &conditional_format)?;

        let conditional_format =
            ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::FourArrows);
        worksheet.add_conditional_format(4, 0, 4, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::FourRedToBlack);
        worksheet.add_conditional_format(5, 0, 5, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::FourTrafficLights);
        worksheet.add_conditional_format(6, 0, 6, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::FiveArrowsGray);
        worksheet.add_conditional_format(7, 0, 7, 0, &conditional_format)?;

        let conditional_format =
            ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::FiveQuadrants);
        worksheet.add_conditional_format(8, 0, 8, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:A9"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:1">
                  <c r="A1">
                    <v>1</v>
                  </c>
                </row>
                <row r="2" spans="1:1">
                  <c r="A2">
                    <v>2</v>
                  </c>
                </row>
                <row r="3" spans="1:1">
                  <c r="A3">
                    <v>3</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="A4">
                    <v>4</v>
                  </c>
                </row>
                <row r="5" spans="1:1">
                  <c r="A5">
                    <v>5</v>
                  </c>
                </row>
                <row r="6" spans="1:1">
                  <c r="A6">
                    <v>6</v>
                  </c>
                </row>
                <row r="7" spans="1:1">
                  <c r="A7">
                    <v>7</v>
                  </c>
                </row>
                <row r="8" spans="1:1">
                  <c r="A8">
                    <v>8</v>
                  </c>
                </row>
                <row r="9" spans="1:1">
                  <c r="A9">
                    <v>9</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A1">
                <cfRule type="iconSet" priority="1">
                  <iconSet iconSet="3Arrows">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="33"/>
                    <cfvo type="percent" val="67"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A2">
                <cfRule type="iconSet" priority="2">
                  <iconSet iconSet="3Flags">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="33"/>
                    <cfvo type="percent" val="67"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A3">
                <cfRule type="iconSet" priority="3">
                  <iconSet iconSet="3TrafficLights2">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="33"/>
                    <cfvo type="percent" val="67"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A4">
                <cfRule type="iconSet" priority="4">
                  <iconSet iconSet="3Symbols">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="33"/>
                    <cfvo type="percent" val="67"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A5">
                <cfRule type="iconSet" priority="5">
                  <iconSet iconSet="4Arrows">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="25"/>
                    <cfvo type="percent" val="50"/>
                    <cfvo type="percent" val="75"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A6">
                <cfRule type="iconSet" priority="6">
                  <iconSet iconSet="4RedToBlack">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="25"/>
                    <cfvo type="percent" val="50"/>
                    <cfvo type="percent" val="75"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A7">
                <cfRule type="iconSet" priority="7">
                  <iconSet iconSet="4TrafficLights">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="25"/>
                    <cfvo type="percent" val="50"/>
                    <cfvo type="percent" val="75"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A8">
                <cfRule type="iconSet" priority="8">
                  <iconSet iconSet="5ArrowsGray">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="20"/>
                    <cfvo type="percent" val="40"/>
                    <cfvo type="percent" val="60"/>
                    <cfvo type="percent" val="80"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A9">
                <cfRule type="iconSet" priority="9">
                  <iconSet iconSet="5Quarters">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="20"/>
                    <cfvo type="percent" val="40"/>
                    <cfvo type="percent" val="60"/>
                    <cfvo type="percent" val="80"/>
                  </iconSet>
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
    fn conditional_format_23() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 1)?;
        worksheet.write(1, 0, 2)?;
        worksheet.write(2, 0, 3)?;
        worksheet.write(3, 0, 4)?;
        worksheet.write(4, 0, 5)?;
        worksheet.write(5, 0, 6)?;
        worksheet.write(6, 0, 7)?;
        worksheet.write(7, 0, 8)?;

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeArrowsGray);
        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights);
        worksheet.add_conditional_format(1, 0, 1, 0, &conditional_format)?;

        let conditional_format =
            ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::ThreeSigns);
        worksheet.add_conditional_format(2, 0, 2, 0, &conditional_format)?;

        let conditional_format =
            ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::ThreeSymbols);
        worksheet.add_conditional_format(3, 0, 3, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::FourArrowsGray);
        worksheet.add_conditional_format(4, 0, 4, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::FourHistograms);
        worksheet.add_conditional_format(5, 0, 5, 0, &conditional_format)?;

        let conditional_format =
            ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::FiveArrows);
        worksheet.add_conditional_format(6, 0, 6, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::FiveHistograms);
        worksheet.add_conditional_format(7, 0, 7, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:A8"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:1">
                  <c r="A1">
                    <v>1</v>
                  </c>
                </row>
                <row r="2" spans="1:1">
                  <c r="A2">
                    <v>2</v>
                  </c>
                </row>
                <row r="3" spans="1:1">
                  <c r="A3">
                    <v>3</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="A4">
                    <v>4</v>
                  </c>
                </row>
                <row r="5" spans="1:1">
                  <c r="A5">
                    <v>5</v>
                  </c>
                </row>
                <row r="6" spans="1:1">
                  <c r="A6">
                    <v>6</v>
                  </c>
                </row>
                <row r="7" spans="1:1">
                  <c r="A7">
                    <v>7</v>
                  </c>
                </row>
                <row r="8" spans="1:1">
                  <c r="A8">
                    <v>8</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A1">
                <cfRule type="iconSet" priority="1">
                  <iconSet iconSet="3ArrowsGray">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="33"/>
                    <cfvo type="percent" val="67"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A2">
                <cfRule type="iconSet" priority="2">
                  <iconSet>
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="33"/>
                    <cfvo type="percent" val="67"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A3">
                <cfRule type="iconSet" priority="3">
                  <iconSet iconSet="3Signs">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="33"/>
                    <cfvo type="percent" val="67"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A4">
                <cfRule type="iconSet" priority="4">
                  <iconSet iconSet="3Symbols2">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="33"/>
                    <cfvo type="percent" val="67"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A5">
                <cfRule type="iconSet" priority="5">
                  <iconSet iconSet="4ArrowsGray">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="25"/>
                    <cfvo type="percent" val="50"/>
                    <cfvo type="percent" val="75"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A6">
                <cfRule type="iconSet" priority="6">
                  <iconSet iconSet="4Rating">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="25"/>
                    <cfvo type="percent" val="50"/>
                    <cfvo type="percent" val="75"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A7">
                <cfRule type="iconSet" priority="7">
                  <iconSet iconSet="5Arrows">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="20"/>
                    <cfvo type="percent" val="40"/>
                    <cfvo type="percent" val="60"/>
                    <cfvo type="percent" val="80"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A8">
                <cfRule type="iconSet" priority="8">
                  <iconSet iconSet="5Rating">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="20"/>
                    <cfvo type="percent" val="40"/>
                    <cfvo type="percent" val="60"/>
                    <cfvo type="percent" val="80"/>
                  </iconSet>
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
    fn conditional_format_24() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, 1)?;
        worksheet.write(1, 0, 2)?;
        worksheet.write(2, 0, 3)?;
        worksheet.write(3, 0, 4)?;
        worksheet.write(4, 0, 5)?;
        worksheet.write(5, 0, 6)?;
        worksheet.write(6, 0, 7)?;
        worksheet.write(7, 0, 8)?;
        worksheet.write(8, 0, 9)?;

        worksheet.write(11, 0, 75)?;

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeArrows)
            .reverse_icons(true);
        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeFlags)
            .show_icons_only(true);
        worksheet.add_conditional_format(1, 0, 1, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeTrafficLightsWithRim)
            .reverse_icons(true)
            .show_icons_only(true);
        worksheet.add_conditional_format(2, 0, 2, 0, &conditional_format)?;

        let icons = [
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 20),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 80),
        ];
        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeSymbolsCircled)
            .set_icons(&icons);
        worksheet.add_conditional_format(3, 0, 3, 0, &conditional_format)?;

        let icons = [
            ConditionalFormatCustomIcon::new()
                .set_rule(ConditionalFormatType::Percentile, 99)
                .set_greater_than(true), // Should be overridden by default.
            ConditionalFormatCustomIcon::new()
                .set_rule(ConditionalFormatType::Percent, 25)
                .set_greater_than(true),
            ConditionalFormatCustomIcon::new()
                .set_rule(ConditionalFormatType::Percent, 50)
                .set_greater_than(true),
            ConditionalFormatCustomIcon::new()
                .set_rule(ConditionalFormatType::Percent, 75)
                .set_greater_than(true),
        ];
        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::FourArrows)
            .set_icons(&icons);
        worksheet.add_conditional_format(4, 0, 4, 0, &conditional_format)?;

        let icons = [
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 25),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percentile, 50),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Number, 90),
        ];
        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::FourRedToBlack)
            .set_icons(&icons);
        worksheet.add_conditional_format(5, 0, 5, 0, &conditional_format)?;

        let icons = [
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 25),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 50),
            ConditionalFormatCustomIcon::new()
                .set_rule(ConditionalFormatType::Percent, Formula::new("$A$12")),
        ];
        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::FourTrafficLights)
            .set_icons(&icons);
        worksheet.add_conditional_format(6, 0, 6, 0, &conditional_format)?;

        let icons = [
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 20),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 40),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 60),
            ConditionalFormatCustomIcon::new()
                .set_rule(ConditionalFormatType::Formula, Formula::new("$A$12")),
        ];
        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::FiveArrowsGray)
            .set_icons(&icons);
        worksheet.add_conditional_format(7, 0, 7, 0, &conditional_format)?;

        let icons = [
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percentile, 10),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percentile, 30),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percentile, 50),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percentile, 70),
        ];
        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::FiveQuadrants)
            .reverse_icons(true)
            .set_icons(&icons);
        worksheet.add_conditional_format(8, 0, 8, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1:A12"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData>
                <row r="1" spans="1:1">
                  <c r="A1">
                    <v>1</v>
                  </c>
                </row>
                <row r="2" spans="1:1">
                  <c r="A2">
                    <v>2</v>
                  </c>
                </row>
                <row r="3" spans="1:1">
                  <c r="A3">
                    <v>3</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="A4">
                    <v>4</v>
                  </c>
                </row>
                <row r="5" spans="1:1">
                  <c r="A5">
                    <v>5</v>
                  </c>
                </row>
                <row r="6" spans="1:1">
                  <c r="A6">
                    <v>6</v>
                  </c>
                </row>
                <row r="7" spans="1:1">
                  <c r="A7">
                    <v>7</v>
                  </c>
                </row>
                <row r="8" spans="1:1">
                  <c r="A8">
                    <v>8</v>
                  </c>
                </row>
                <row r="9" spans="1:1">
                  <c r="A9">
                    <v>9</v>
                  </c>
                </row>
                <row r="12" spans="1:1">
                  <c r="A12">
                    <v>75</v>
                  </c>
                </row>
              </sheetData>
              <conditionalFormatting sqref="A1">
                <cfRule type="iconSet" priority="1">
                  <iconSet iconSet="3Arrows" reverse="1">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="33"/>
                    <cfvo type="percent" val="67"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A2">
                <cfRule type="iconSet" priority="2">
                  <iconSet iconSet="3Flags" showValue="0">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="33"/>
                    <cfvo type="percent" val="67"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A3">
                <cfRule type="iconSet" priority="3">
                  <iconSet iconSet="3TrafficLights2" showValue="0" reverse="1">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="33"/>
                    <cfvo type="percent" val="67"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A4">
                <cfRule type="iconSet" priority="4">
                  <iconSet iconSet="3Symbols">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="20"/>
                    <cfvo type="percent" val="80"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A5">
                <cfRule type="iconSet" priority="5">
                  <iconSet iconSet="4Arrows">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="25" gte="0"/>
                    <cfvo type="percent" val="50" gte="0"/>
                    <cfvo type="percent" val="75" gte="0"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A6">
                <cfRule type="iconSet" priority="6">
                  <iconSet iconSet="4RedToBlack">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="25"/>
                    <cfvo type="percentile" val="50"/>
                    <cfvo type="num" val="90"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A7">
                <cfRule type="iconSet" priority="7">
                  <iconSet iconSet="4TrafficLights">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="25"/>
                    <cfvo type="percent" val="50"/>
                    <cfvo type="percent" val="$A$12"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A8">
                <cfRule type="iconSet" priority="8">
                  <iconSet iconSet="5ArrowsGray">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="20"/>
                    <cfvo type="percent" val="40"/>
                    <cfvo type="percent" val="60"/>
                    <cfvo type="formula" val="$A$12"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A9">
                <cfRule type="iconSet" priority="9">
                  <iconSet iconSet="5Quarters" reverse="1">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percentile" val="10"/>
                    <cfvo type="percentile" val="30"/>
                    <cfvo type="percentile" val="50"/>
                    <cfvo type="percentile" val="70"/>
                  </iconSet>
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
    fn data_bar_01() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatDataBar::new().use_classic_style();

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
                <cfRule type="dataBar" priority="1">
                  <dataBar>
                    <cfvo type="min" val="0"/>
                    <cfvo type="max" val="0"/>
                    <color rgb="FF638EC6"/>
                  </dataBar>
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
    fn data_bar_02() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatDataBar::new();

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
                <cfRule type="dataBar" priority="1">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF638EC6"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000001}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FF638EC6"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_bar_03() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatDataBar::new();
        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new().set_fill_color("63C384");
        worksheet.add_conditional_format(1, 0, 1, 1, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new().set_fill_color("FF555A");
        worksheet.add_conditional_format(2, 0, 2, 2, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
                <cfRule type="dataBar" priority="1">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF638EC6"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000001}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A2:B2">
                <cfRule type="dataBar" priority="2">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF63C384"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000002}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A3:C3">
                <cfRule type="dataBar" priority="3">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FFFF555A"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000003}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FF638EC6"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000002}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FF63C384"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A2:B2</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000003}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FFFF555A"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A3:C3</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_bar_04() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatDataBar::new().set_solid_fill(true);
        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_fill_color("63C384")
            .set_border_off(true);
        worksheet.add_conditional_format(1, 0, 1, 1, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_fill_color("FF555A")
            .set_border_color("FF0000");
        worksheet.add_conditional_format(2, 0, 2, 2, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
                <cfRule type="dataBar" priority="1">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF638EC6"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000001}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A2:B2">
                <cfRule type="dataBar" priority="2">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF63C384"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000002}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A3:C3">
                <cfRule type="dataBar" priority="3">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FFFF555A"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000003}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" gradient="0" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FF638EC6"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000002}">
                        <x14:dataBar minLength="0" maxLength="100">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A2:B2</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000003}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FFFF0000"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A3:C3</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_bar_05() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatDataBar::new()
            .set_direction(ConditionalFormatDataBarDirection::LeftToRight);
        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_fill_color("63C384")
            .set_direction(ConditionalFormatDataBarDirection::RightToLeft);
        worksheet.add_conditional_format(1, 0, 1, 1, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_fill_color("FF555A")
            .set_negative_fill_color("FFFF00")
            .set_negative_border_color("FF0000");
        worksheet.add_conditional_format(2, 0, 2, 2, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
                <cfRule type="dataBar" priority="1">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF638EC6"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000001}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A2:B2">
                <cfRule type="dataBar" priority="2">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF63C384"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000002}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A3:C3">
                <cfRule type="dataBar" priority="3">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FFFF555A"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000003}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" direction="leftToRight" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FF638EC6"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000002}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" direction="rightToLeft" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FF63C384"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A2:B2</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000003}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FFFF555A"/>
                          <x14:negativeFillColor rgb="FFFFFF00"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A3:C3</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_bar_06() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatDataBar::new()
            .set_negative_fill_color("638EC6")
            .set_negative_border_color("FF0000");
        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_fill_color("63C384")
            .set_negative_border_color("92D050");
        worksheet.add_conditional_format(1, 0, 1, 1, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_fill_color("FF555A")
            .set_negative_border_color("FF555A");
        worksheet.add_conditional_format(2, 0, 2, 2, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
                <cfRule type="dataBar" priority="1">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF638EC6"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000001}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A2:B2">
                <cfRule type="dataBar" priority="2">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF63C384"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000002}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A3:C3">
                <cfRule type="dataBar" priority="3">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FFFF555A"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000003}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarColorSameAsPositive="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FF638EC6"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000002}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FF63C384"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FF92D050"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A2:B2</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000003}">
                        <x14:dataBar minLength="0" maxLength="100" border="1">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FFFF555A"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A3:C3</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_bar_07() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatDataBar::new()
            .set_axis_position(ConditionalFormatDataBarAxisPosition::Midpoint);
        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_fill_color("63C384")
            .set_axis_position(ConditionalFormatDataBarAxisPosition::None);
        worksheet.add_conditional_format(1, 0, 1, 1, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_fill_color("FF555A")
            .set_axis_color("0070C0");
        worksheet.add_conditional_format(2, 0, 2, 2, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
                <cfRule type="dataBar" priority="1">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF638EC6"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000001}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A2:B2">
                <cfRule type="dataBar" priority="2">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF63C384"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000002}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A3:C3">
                <cfRule type="dataBar" priority="3">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FFFF555A"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000003}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0" axisPosition="middle">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FF638EC6"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000002}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0" axisPosition="none">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FF63C384"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A2:B2</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000003}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FFFF555A"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF0070C0"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A3:C3</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_bar_08() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatDataBar::new()
            .use_classic_style()
            .set_bar_only(true);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
                <cfRule type="dataBar" priority="1">
                  <dataBar showValue="0">
                    <cfvo type="min" val="0"/>
                    <cfvo type="max" val="0"/>
                    <color rgb="FF638EC6"/>
                  </dataBar>
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
    fn data_bar_09() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatDataBar::new().set_bar_only(true);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
                <cfRule type="dataBar" priority="1">
                  <dataBar showValue="0">
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF638EC6"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000001}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FF638EC6"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_bar_10() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatDataBar::new()
            .set_minimum(ConditionalFormatType::Lowest, 98)
            .set_maximum(ConditionalFormatType::Highest, 99);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_fill_color("63C384")
            .set_minimum(ConditionalFormatType::Number, 0)
            .set_maximum(ConditionalFormatType::Number, 0);

        worksheet.add_conditional_format(1, 0, 1, 1, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_fill_color("FF555A")
            .set_minimum(ConditionalFormatType::Percent, 0)
            .set_maximum(ConditionalFormatType::Percent, 100);

        worksheet.add_conditional_format(2, 0, 2, 2, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
                <cfRule type="dataBar" priority="1">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF638EC6"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000001}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A2:B2">
                <cfRule type="dataBar" priority="2">
                  <dataBar>
                    <cfvo type="num" val="0"/>
                    <cfvo type="num" val="0"/>
                    <color rgb="FF63C384"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000002}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A3:C3">
                <cfRule type="dataBar" priority="3">
                  <dataBar>
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="100"/>
                    <color rgb="FFFF555A"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000003}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="min"/>
                          <x14:cfvo type="max"/>
                          <x14:borderColor rgb="FF638EC6"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000002}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="num">
                            <xm:f>0</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="num">
                            <xm:f>0</xm:f>
                          </x14:cfvo>
                          <x14:borderColor rgb="FF63C384"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A2:B2</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000003}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="percent">
                            <xm:f>0</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>100</xm:f>
                          </x14:cfvo>
                          <x14:borderColor rgb="FFFF555A"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A3:C3</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_bar_11() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatDataBar::new()
            .set_minimum(ConditionalFormatType::Formula, Formula::new("=$B$1"));

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_fill_color("63C384")
            .set_minimum(ConditionalFormatType::Formula, Formula::new("=$B$1"))
            .set_maximum(ConditionalFormatType::Formula, Formula::new("=$C$1"));

        worksheet.add_conditional_format(1, 0, 1, 1, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_fill_color("FF555A")
            .set_minimum(ConditionalFormatType::Percentile, 10)
            .set_maximum(ConditionalFormatType::Percentile, 90);

        worksheet.add_conditional_format(2, 0, 2, 2, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
                <cfRule type="dataBar" priority="1">
                  <dataBar>
                    <cfvo type="formula" val="$B$1"/>
                    <cfvo type="max"/>
                    <color rgb="FF638EC6"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000001}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A2:B2">
                <cfRule type="dataBar" priority="2">
                  <dataBar>
                    <cfvo type="formula" val="$B$1"/>
                    <cfvo type="formula" val="$C$1"/>
                    <color rgb="FF63C384"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000002}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <conditionalFormatting sqref="A3:C3">
                <cfRule type="dataBar" priority="3">
                  <dataBar>
                    <cfvo type="percentile" val="10"/>
                    <cfvo type="percentile" val="90"/>
                    <color rgb="FFFF555A"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000003}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="formula">
                            <xm:f>$B$1</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FF638EC6"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000002}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="formula">
                            <xm:f>$B$1</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="formula">
                            <xm:f>$C$1</xm:f>
                          </x14:cfvo>
                          <x14:borderColor rgb="FF63C384"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A2:B2</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000003}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="percentile">
                            <xm:f>10</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percentile">
                            <xm:f>90</xm:f>
                          </x14:cfvo>
                          <x14:borderColor rgb="FFFF555A"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A3:C3</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_bar_13() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatDataBar::new();

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
                <cfRule type="dataBar" priority="1">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF638EC6"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000001}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:borderColor rgb="FF638EC6"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:negativeBorderColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn icon_01() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
                <cfRule type="iconSet" priority="1">
                  <iconSet>
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="33"/>
                    <cfvo type="percent" val="67"/>
                  </iconSet>
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
    fn icon_02() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format =
            ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::ThreeStars);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="iconSet" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:iconSet iconSet="3Stars">
                          <x14:cfvo type="percent">
                            <xm:f>0</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>33</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>67</xm:f>
                          </x14:cfvo>
                        </x14:iconSet>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn icon_03() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeTriangles);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="iconSet" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:iconSet iconSet="3Triangles">
                          <x14:cfvo type="percent">
                            <xm:f>0</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>33</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>67</xm:f>
                          </x14:cfvo>
                        </x14:iconSet>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn icon_04() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format =
            ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::FiveBoxes);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="iconSet" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:iconSet iconSet="5Boxes">
                        <x14:cfvo type="percent">
                          <xm:f>0</xm:f>
                        </x14:cfvo>
                        <x14:cfvo type="percent">
                          <xm:f>20</xm:f>
                        </x14:cfvo>
                        <x14:cfvo type="percent">
                          <xm:f>40</xm:f>
                        </x14:cfvo>
                        <x14:cfvo type="percent">
                          <xm:f>60</xm:f>
                        </x14:cfvo>
                        <x14:cfvo type="percent">
                          <xm:f>80</xm:f>
                        </x14:cfvo>
                      </x14:iconSet>
                    </x14:cfRule>
                    <xm:sqref>A1</xm:sqref>
                  </x14:conditionalFormatting>
                </x14:conditionalFormattings>
              </ext>
            </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn icon_05() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format =
            ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::ThreeStars);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
              <cfRule type="iconSet" priority="1">
                <iconSet>
                  <cfvo type="percent" val="0"/>
                  <cfvo type="percent" val="33"/>
                  <cfvo type="percent" val="67"/>
                </iconSet>
              </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="iconSet" priority="2" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:iconSet iconSet="3Stars">
                          <x14:cfvo type="percent">
                            <xm:f>0</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>33</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>67</xm:f>
                          </x14:cfvo>
                        </x14:iconSet>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn icon_06() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format =
            ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::ThreeStars);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <conditionalFormatting sqref="A1">
                <cfRule type="iconSet" priority="2">
                  <iconSet>
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="33"/>
                    <cfvo type="percent" val="67"/>
                  </iconSet>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="iconSet" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:iconSet iconSet="3Stars">
                          <x14:cfvo type="percent">
                            <xm:f>0</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>33</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>67</xm:f>
                          </x14:cfvo>
                        </x14:iconSet>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);
        Ok(())
    }

    #[test]
    fn icon_07() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format =
            ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::ThreeStars);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_solid_fill(true)
            .set_border_off(true);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
            <sheetData/>
            <conditionalFormatting sqref="A1">
              <cfRule type="dataBar" priority="2">
                <dataBar>
                  <cfvo type="min"/>
                  <cfvo type="max"/>
                  <color rgb="FF638EC6"/>
                </dataBar>
                <extLst>
                  <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                    <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000002}</x14:id>
                  </ext>
                </extLst>
              </cfRule>
            </conditionalFormatting>
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            <extLst>
              <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                <x14:conditionalFormattings>
                  <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:cfRule type="iconSet" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                      <x14:iconSet iconSet="3Stars">
                        <x14:cfvo type="percent">
                          <xm:f>0</xm:f>
                        </x14:cfvo>
                        <x14:cfvo type="percent">
                          <xm:f>33</xm:f>
                        </x14:cfvo>
                        <x14:cfvo type="percent">
                          <xm:f>67</xm:f>
                        </x14:cfvo>
                      </x14:iconSet>
                    </x14:cfRule>
                    <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000002}">
                      <x14:dataBar minLength="0" maxLength="100" gradient="0">
                        <x14:cfvo type="autoMin"/>
                        <x14:cfvo type="autoMax"/>
                        <x14:negativeFillColor rgb="FFFF0000"/>
                        <x14:axisColor rgb="FF000000"/>
                      </x14:dataBar>
                    </x14:cfRule>
                    <xm:sqref>A1</xm:sqref>
                  </x14:conditionalFormatting>
                </x14:conditionalFormattings>
              </ext>
            </extLst>
          </worksheet>
          "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn icon_08() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format =
            ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::ThreeStars);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_solid_fill(true)
            .set_border_off(true);

        worksheet.add_conditional_format(0, 1, 0, 1, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
              <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
              <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
                <dimension ref="A1"/>
                <sheetViews>
                  <sheetView tabSelected="1" workbookViewId="0"/>
                </sheetViews>
                <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <conditionalFormatting sqref="B1">
                <cfRule type="dataBar" priority="2">
                  <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF638EC6"/>
                  </dataBar>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                      <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000002}</x14:id>
                    </ext>
                  </extLst>
                </cfRule>
              </conditionalFormatting>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="iconSet" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:iconSet iconSet="3Stars">
                          <x14:cfvo type="percent">
                            <xm:f>0</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>33</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>67</xm:f>
                          </x14:cfvo>
                        </x14:iconSet>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000002}">
                        <x14:dataBar minLength="0" maxLength="100" gradient="0">
                          <x14:cfvo type="autoMin"/>
                          <x14:cfvo type="autoMax"/>
                          <x14:negativeFillColor rgb="FFFF0000"/>
                          <x14:axisColor rgb="FF000000"/>
                        </x14:dataBar>
                      </x14:cfRule>
                      <xm:sqref>B1</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn icon_09() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format =
            ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::ThreeStars);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        let conditional_format = ConditionalFormatDataBar::new()
            .set_solid_fill(true)
            .set_border_off(true);

        worksheet.add_conditional_format(0, 1, 0, 1, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
            <sheetData/>
            <conditionalFormatting sqref="A1">
              <cfRule type="iconSet" priority="1">
                <iconSet>
                  <cfvo type="percent" val="0"/>
                  <cfvo type="percent" val="33"/>
                  <cfvo type="percent" val="67"/>
                </iconSet>
              </cfRule>
            </conditionalFormatting>
            <conditionalFormatting sqref="B1">
              <cfRule type="dataBar" priority="3">
                <dataBar>
                  <cfvo type="min"/>
                  <cfvo type="max"/>
                  <color rgb="FF638EC6"/>
                </dataBar>
                <extLst>
                  <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                    <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000002}</x14:id>
                  </ext>
                </extLst>
              </cfRule>
            </conditionalFormatting>
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            <extLst>
              <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                <x14:conditionalFormattings>
                  <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:cfRule type="iconSet" priority="2" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                      <x14:iconSet iconSet="3Stars">
                        <x14:cfvo type="percent">
                          <xm:f>0</xm:f>
                        </x14:cfvo>
                        <x14:cfvo type="percent">
                          <xm:f>33</xm:f>
                        </x14:cfvo>
                        <x14:cfvo type="percent">
                          <xm:f>67</xm:f>
                        </x14:cfvo>
                      </x14:iconSet>
                    </x14:cfRule>
                    <xm:sqref>A1</xm:sqref>
                  </x14:conditionalFormatting>
                  <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000002}">
                      <x14:dataBar minLength="0" maxLength="100" gradient="0">
                        <x14:cfvo type="autoMin"/>
                        <x14:cfvo type="autoMax"/>
                        <x14:negativeFillColor rgb="FFFF0000"/>
                        <x14:axisColor rgb="FF000000"/>
                      </x14:dataBar>
                    </x14:cfRule>
                    <xm:sqref>B1</xm:sqref>
                  </x14:conditionalFormatting>
                </x14:conditionalFormattings>
              </ext>
            </extLst>
          </worksheet>
          "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn icon_10() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let icons = [
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 33),
            ConditionalFormatCustomIcon::new()
                .set_rule(ConditionalFormatType::Percent, 67)
                .set_icon_type(ConditionalFormatIconType::ThreeArrows, 2),
        ];
        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeStars)
            .set_icons(&icons);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
              <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
              <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
                <dimension ref="A1"/>
                <sheetViews>
                  <sheetView tabSelected="1" workbookViewId="0"/>
                </sheetViews>
                <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="iconSet" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:iconSet iconSet="3Stars" custom="1">
                          <x14:cfvo type="percent">
                            <xm:f>0</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>33</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>67</xm:f>
                          </x14:cfvo>
                          <x14:cfIcon iconSet="3Stars" iconId="0"/>
                          <x14:cfIcon iconSet="3Stars" iconId="1"/>
                          <x14:cfIcon iconSet="3Arrows" iconId="2"/>
                        </x14:iconSet>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn icon_11() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let icons = [
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 33),
            ConditionalFormatCustomIcon::new()
                .set_rule(ConditionalFormatType::Percent, 67)
                .set_icon_type(ConditionalFormatIconType::FiveBoxes, 0),
        ];
        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeStars)
            .set_icons(&icons);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
              <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
              <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
                <dimension ref="A1"/>
                <sheetViews>
                  <sheetView tabSelected="1" workbookViewId="0"/>
                </sheetViews>
                <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="iconSet" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:iconSet iconSet="3Stars" custom="1">
                          <x14:cfvo type="percent">
                            <xm:f>0</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>33</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>67</xm:f>
                          </x14:cfvo>
                          <x14:cfIcon iconSet="3Stars" iconId="0"/>
                          <x14:cfIcon iconSet="3Stars" iconId="1"/>
                          <x14:cfIcon iconSet="5Boxes" iconId="0"/>
                        </x14:iconSet>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn icon_12() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let icons = [
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 33),
            ConditionalFormatCustomIcon::new()
                .set_rule(ConditionalFormatType::Percent, 67)
                .set_icon_type(ConditionalFormatIconType::FiveBoxes, 4),
        ];
        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeStars)
            .set_icons(&icons);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
              <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
              <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
                <dimension ref="A1"/>
                <sheetViews>
                  <sheetView tabSelected="1" workbookViewId="0"/>
                </sheetViews>
                <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="iconSet" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:iconSet iconSet="3Stars" custom="1">
                          <x14:cfvo type="percent">
                            <xm:f>0</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>33</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>67</xm:f>
                          </x14:cfvo>
                          <x14:cfIcon iconSet="3Stars" iconId="0"/>
                          <x14:cfIcon iconSet="3Stars" iconId="1"/>
                          <x14:cfIcon iconSet="5Boxes" iconId="4"/>
                        </x14:iconSet>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn icon_13() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let icons = [
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 33),
            ConditionalFormatCustomIcon::new()
                .set_rule(ConditionalFormatType::Percent, 67)
                .set_icon_type(ConditionalFormatIconType::FiveHistograms, 0),
        ];
        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights)
            .set_icons(&icons);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
              <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
              <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
                <dimension ref="A1"/>
                <sheetViews>
                  <sheetView tabSelected="1" workbookViewId="0"/>
                </sheetViews>
                <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="iconSet" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:iconSet custom="1">
                          <x14:cfvo type="percent">
                            <xm:f>0</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>33</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>67</xm:f>
                          </x14:cfvo>
                          <x14:cfIcon iconSet="3TrafficLights1" iconId="0"/>
                          <x14:cfIcon iconSet="3TrafficLights1" iconId="1"/>
                          <x14:cfIcon iconSet="5Rating" iconId="0"/>
                        </x14:iconSet>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn icon_14() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let icons = [
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
            ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 33),
            ConditionalFormatCustomIcon::new()
                .set_rule(ConditionalFormatType::Percent, 67)
                .set_no_icon(true),
        ];
        let conditional_format = ConditionalFormatIconSet::new()
            .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights)
            .set_icons(&icons);

        worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
              <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
              <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
                <dimension ref="A1"/>
                <sheetViews>
                  <sheetView tabSelected="1" workbookViewId="0"/>
                </sheetViews>
                <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData/>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                  <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="iconSet" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                        <x14:iconSet custom="1">
                          <x14:cfvo type="percent">
                            <xm:f>0</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>33</xm:f>
                          </x14:cfvo>
                          <x14:cfvo type="percent">
                            <xm:f>67</xm:f>
                          </x14:cfvo>
                          <x14:cfIcon iconSet="3TrafficLights1" iconId="0"/>
                          <x14:cfIcon iconSet="3TrafficLights1" iconId="1"/>
                          <x14:cfIcon iconSet="NoIcons" iconId="0"/>
                        </x14:iconSet>
                      </x14:cfRule>
                      <xm:sqref>A1</xm:sqref>
                    </x14:conditionalFormatting>
                  </x14:conditionalFormattings>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }
}
