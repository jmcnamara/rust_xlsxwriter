// data_validation unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod data_validation_tests {

    use crate::test_functions::xml_to_vec;
    use crate::DataValidation;
    use crate::DataValidationErrorStyle;
    use crate::DataValidationRule;
    use crate::DataValidationType;
    use crate::ExcelDateTime;
    use crate::Worksheet;
    use crate::XlsxError;

    use pretty_assertions::assert_eq;

    #[test]
    fn data_validation_01() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Whole)
            .set_rule(DataValidationRule::GreaterThan(0));

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>0</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_02() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Decimal)
            .set_rule(DataValidationRule::Between(1.0, 2.0));

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="decimal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>1</formula1>
                  <formula2>2</formula2>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_03() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Whole)
            .set_rule(DataValidationRule::LessThan(10))
            .ignore_blank(false);

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="whole" operator="lessThan" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>10</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_04() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Whole)
            .set_rule(DataValidationRule::LessThan(10))
            .show_input_message(false);

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="whole" operator="lessThan" allowBlank="1" showErrorMessage="1" sqref="A1">
                  <formula1>10</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_05() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Whole)
            .set_rule(DataValidationRule::LessThan(10))
            .show_error_message(false);

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="whole" operator="lessThan" allowBlank="1" showInputMessage="1" sqref="A1">
                  <formula1>10</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_06() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Whole)
            .set_rule(DataValidationRule::LessThan(10))
            .show_error_message(false);

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="whole" operator="lessThan" allowBlank="1" showInputMessage="1" sqref="A1">
                  <formula1>10</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_07() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Whole)
            .set_rule(DataValidationRule::NotEqualTo(10))
            .set_input_title("Title 1");

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="whole" operator="notEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Title 1" sqref="A1">
                  <formula1>10</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_08() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Whole)
            .set_rule(DataValidationRule::NotEqualTo(10))
            .set_input_title("Title 1")
            .set_input_message("Message 1");

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="whole" operator="notEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Title 1" prompt="Message 1" sqref="A1">
                  <formula1>10</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_09() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Whole)
            .set_rule(DataValidationRule::NotEqualTo(10))
            .set_error_title("Title 2")
            .set_error_message("Message 2")
            .set_input_title("Title 1")
            .set_input_message("Message 1");

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="whole" operator="notEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Title 2" error="Message 2" promptTitle="Title 1" prompt="Message 1" sqref="A1">
                  <formula1>10</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_10() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Whole)
            .set_rule(DataValidationRule::NotEqualTo(10))
            .set_error_style(DataValidationErrorStyle::Warning);

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="whole" errorStyle="warning" operator="notEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>10</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_11() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Whole)
            .set_rule(DataValidationRule::NotEqualTo(10))
            .set_error_style(DataValidationErrorStyle::Information);

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="whole" errorStyle="information" operator="notEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>10</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_12_1() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Date)
            .set_rule(DataValidationRule::GreaterThan(45658));

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="date" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>45658</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_12_2() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Date)
            .set_rule(DataValidationRule::GreaterThan(
                ExcelDateTime::parse_from_str("2025-01-01")?,
            ));

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="date" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>45658</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_13_1() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Time)
            .set_rule(DataValidationRule::GreaterThan(0.5));

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="time" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>0.5</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_13_2() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Time)
            .set_rule(DataValidationRule::GreaterThan(
                ExcelDateTime::parse_from_str("12:00")?,
            ));

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="time" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>0.5</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_14() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::TextLength)
            .set_rule(DataValidationRule::GreaterThan(6));

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="textLength" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>6</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_15() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::Custom)
            .set_rule(DataValidationRule::CustomFormula(6));

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>6</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_16_1() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::List)
            .set_rule(DataValidationRule::ListSource("\"Foo,Bar,Baz\""));

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>"Foo,Bar,Baz"</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_16_2() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::List)
            .set_string_list(&["Foo", "Bar", "Baz"]);

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>"Foo,Bar,Baz"</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_16_3() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .set_type(DataValidationType::List)
            .set_string_list(&[
                String::from("Foo"),
                String::from("Bar"),
                String::from("Baz"),
            ]);

        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
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
              <dataValidations count="1">
                <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>"Foo,Bar,Baz"</formula1>
                </dataValidation>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }
}
