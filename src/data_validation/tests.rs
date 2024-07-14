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
    use crate::ExcelDateTime;
    use crate::Formula;
    use crate::Worksheet;
    use crate::XlsxError;

    #[cfg(feature = "chrono")]
    use chrono::{NaiveDate, NaiveTime};

    use pretty_assertions::assert_eq;

    #[test]
    fn data_validation_01() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation =
            DataValidation::new().allow_whole_number(DataValidationRule::GreaterThan(-10));

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
                  <formula1>-10</formula1>
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

        let data_validation =
            DataValidation::new().allow_decimal_number(DataValidationRule::Between(1.0, 2.0));

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
            .allow_whole_number(DataValidationRule::LessThan(10))
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
            .allow_whole_number(DataValidationRule::LessThan(10))
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
            .allow_whole_number(DataValidationRule::LessThan(10))
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
            .allow_whole_number(DataValidationRule::LessThan(10))
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
            .allow_whole_number(DataValidationRule::NotEqualTo(10))
            .set_input_title("Title 1")?;

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
            .allow_whole_number(DataValidationRule::NotEqualTo(10))
            .set_input_title("Title 1")?
            .set_input_message("Message 1")?;

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
            .allow_whole_number(DataValidationRule::NotEqualTo(10))
            .set_error_title("Title 2")?
            .set_error_message("Message 2")?
            .set_input_title("Title 1")?
            .set_input_message("Message 1")?;

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
            .allow_whole_number(DataValidationRule::NotEqualTo(10))
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
            .allow_whole_number(DataValidationRule::NotEqualTo(10))
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

        let data_validation = DataValidation::new().allow_date(DataValidationRule::GreaterThan(
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
    fn data_validation_12_2() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new().allow_date(DataValidationRule::GreaterThan(
            &ExcelDateTime::parse_from_str("2025-01-01")?,
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

    #[cfg(feature = "chrono")]
    #[test]
    fn data_validation_12_3() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new().allow_date(DataValidationRule::GreaterThan(
            NaiveDate::from_ymd_opt(2025, 1, 1).unwrap(),
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

    #[cfg(feature = "chrono")]
    #[test]
    fn data_validation_12_4() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new().allow_date(DataValidationRule::GreaterThan(
            &NaiveDate::from_ymd_opt(2025, 1, 1).unwrap(),
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

    #[cfg(feature = "chrono")]
    #[test]
    fn data_validation_12_5() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new().allow_date(DataValidationRule::GreaterThan(
            NaiveDate::from_ymd_opt(2025, 1, 1)
                .unwrap()
                .and_hms_opt(0, 0, 0)
                .unwrap(),
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

    #[cfg(feature = "chrono")]
    #[test]
    fn data_validation_12_6() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new().allow_date(DataValidationRule::GreaterThan(
            &NaiveDate::from_ymd_opt(2025, 1, 1)
                .unwrap()
                .and_hms_opt(0, 0, 0)
                .unwrap(),
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

        let data_validation = DataValidation::new().allow_time(DataValidationRule::GreaterThan(
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
    fn data_validation_13_2() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new().allow_time(DataValidationRule::GreaterThan(
            &ExcelDateTime::parse_from_str("12:00")?,
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

    #[cfg(feature = "chrono")]
    #[test]
    fn data_validation_13_3() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new().allow_time(DataValidationRule::GreaterThan(
            NaiveTime::from_hms_milli_opt(12, 0, 0, 0).unwrap(),
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

    #[cfg(feature = "chrono")]
    #[test]
    fn data_validation_13_4() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new().allow_time(DataValidationRule::GreaterThan(
            &NaiveTime::from_hms_milli_opt(12, 0, 0, 0).unwrap(),
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

        let data_validation =
            DataValidation::new().allow_text_length(DataValidationRule::GreaterThan(6));

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
    fn data_validation_15_1() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new().allow_custom(Formula::new("=6"));

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
    fn data_validation_15_2() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new().allow_custom("6".into());

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

        let data_validation = DataValidation::new().allow_list_strings(&["Foo", "Bar", "Baz"])?;

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

        let data_validation = DataValidation::new().allow_list_strings(&[
            String::from("Foo"),
            String::from("Bar"),
            String::from("Baz"),
        ])?;

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

        let data_validation = DataValidation::new().allow_list_strings(&[
            &String::from("Foo"),
            &String::from("Bar"),
            &String::from("Baz"),
        ])?;

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
    fn data_validation_17_1() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .allow_any_value()
            .set_input_title("Title 1")?;

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
                <dataValidation allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Title 1" sqref="A1"/>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_17_2() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new().set_input_title("Title 1")?;

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
                <dataValidation allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Title 1" sqref="A1"/>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_18() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .allow_any_value()
            .set_input_message("Message 1")?;

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
                <dataValidation allowBlank="1" showInputMessage="1" showErrorMessage="1" prompt="Message 1" sqref="A1"/>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_19() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .allow_any_value()
            .set_error_title("Title 2")?
            .set_error_message("Message 2")?;

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
                <dataValidation allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Title 2" error="Message 2" sqref="A1"/>
              </dataValidations>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_20() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        // Check for empty "Any" validation.
        let data_validation = DataValidation::new().allow_any_value();

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
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_21() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let invalid_title = "This exceeds Excel's title limits";
        let padding = ["a"; 221];
        let invalid_message = format!("This exceeds Excel's message limits{}", padding.concat());
        let list_values = [
            "Foobar", "Foobas", "Foobat", "Foobau", "Foobav", "Foobaw", "Foobax", "Foobay",
            "Foobaz", "Foobba", "Foobbb", "Foobbc", "Foobbd", "Foobbe", "Foobbf", "Foobbg",
            "Foobbh", "Foobbi", "Foobbj", "Foobbk", "Foobbl", "Foobbm", "Foobbn", "Foobbo",
            "Foobbp", "Foobbq", "Foobbr", "Foobbs", "Foobbt", "Foobbu", "Foobbv", "Foobbw",
            "Foobbx", "Foobby", "Foobbz", "Foobca", "End1",
        ];

        // Check for invalid string lengths.
        let result = DataValidation::new().set_input_title(invalid_title);
        assert!(matches!(result, Err(XlsxError::DataValidationError(_))));

        // Check for invalid string lengths.
        let result = DataValidation::new().set_input_message(&invalid_message);
        assert!(matches!(result, Err(XlsxError::DataValidationError(_))));

        // Check for invalid string lengths.
        let result = DataValidation::new().set_error_title(invalid_title);
        assert!(matches!(result, Err(XlsxError::DataValidationError(_))));

        // Check for invalid string lengths.
        let result = DataValidation::new().set_error_message(&invalid_message);
        assert!(matches!(result, Err(XlsxError::DataValidationError(_))));

        // Check for invalid string list.
        let result = DataValidation::new().allow_list_strings(&list_values);
        assert!(matches!(result, Err(XlsxError::DataValidationError(_))));

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
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn data_validation_22() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .allow_whole_number_formula(DataValidationRule::EqualTo(Formula::new("=J13")));

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
                <dataValidation type="whole" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>J13</formula1>
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
    fn data_validation_23() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .allow_whole_number_formula(DataValidationRule::EqualTo(Formula::new("=$J13")));

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
                <dataValidation type="whole" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>$J13</formula1>
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
    fn data_validation_24() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .allow_whole_number_formula(DataValidationRule::GreaterThan("B1".into()));
        worksheet.add_data_validation(0, 0, 0, 0, &data_validation)?;

        let data_validation = DataValidation::new()
            .allow_decimal_number_formula(DataValidationRule::GreaterThan("B2".into()));
        worksheet.add_data_validation(1, 0, 1, 0, &data_validation)?;

        let data_validation = DataValidation::new().allow_list_formula("$B$3:$E$3".into());
        worksheet.add_data_validation(2, 0, 2, 0, &data_validation)?;

        let data_validation =
            DataValidation::new().allow_date_formula(DataValidationRule::GreaterThan("B4".into()));
        worksheet.add_data_validation(3, 0, 3, 0, &data_validation)?;

        let data_validation =
            DataValidation::new().allow_time_formula(DataValidationRule::GreaterThan("B5".into()));
        worksheet.add_data_validation(4, 0, 4, 0, &data_validation)?;

        let data_validation = DataValidation::new()
            .allow_text_length_formula(DataValidationRule::GreaterThan("B6".into()));
        worksheet.add_data_validation(5, 0, 5, 0, &data_validation)?;

        let data_validation = DataValidation::new().allow_custom("B7".into());
        worksheet.add_data_validation(6, 0, 6, 0, &data_validation)?;

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
              <dataValidations count="7">
                <dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
                  <formula1>B1</formula1>
                </dataValidation>
                <dataValidation type="decimal" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A2">
                  <formula1>B2</formula1>
                </dataValidation>
                <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A3">
                  <formula1>$B$3:$E$3</formula1>
                </dataValidation>
                <dataValidation type="date" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A4">
                  <formula1>B4</formula1>
                </dataValidation>
                <dataValidation type="time" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A5">
                  <formula1>B5</formula1>
                </dataValidation>
                <dataValidation type="textLength" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A6">
                  <formula1>B6</formula1>
                </dataValidation>
                <dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A7">
                  <formula1>B7</formula1>
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
    fn data_validation_25() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation =
            DataValidation::new().allow_whole_number(DataValidationRule::GreaterThan(7));

        worksheet.add_data_validation(0, 0, 5, 0, &data_validation)?;

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
                <dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1:A6">
                  <formula1>7</formula1>
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
    fn data_validation_26() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .allow_whole_number(DataValidationRule::GreaterThan(8))
            .set_multi_range("A1:A6 B8 C10");

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
                <dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1:A6 B8 C10">
                  <formula1>8</formula1>
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
    fn data_validation_27() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data_validation = DataValidation::new()
            .allow_list_strings(&["Foo", "Bar", "Baz"])?
            .show_dropdown(false);

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
                <dataValidation type="list" allowBlank="1" showDropDown="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
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
