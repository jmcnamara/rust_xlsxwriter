// sparkline unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod sparkline_tests {

    use crate::test_functions::xml_to_vec;
    use crate::ChartEmptyCells;
    use crate::Color;
    use crate::Sparkline;
    use crate::SparklineType;
    use crate::Worksheet;
    use crate::XlsxError;
    use pretty_assertions::assert_eq;

    #[test]
    fn sparkline_errors() {
        let mut worksheet = Worksheet::new();

        let sparkline = Sparkline::new();
        let result = worksheet.add_sparkline(0, 5, &sparkline);
        assert!(matches!(result, Err(XlsxError::SparklineError(_))));

        let sparkline = Sparkline::new().set_range(("Sheet1", 0, 0, 1, 4));
        let result = worksheet.add_sparkline(0, 5, &sparkline);
        assert!(matches!(result, Err(XlsxError::SparklineError(_))));

        let sparkline = Sparkline::new().set_range(("Sheet1", 9, 9, 0, 4));
        let result = worksheet.add_sparkline(0, 5, &sparkline);
        assert!(matches!(result, Err(XlsxError::ChartError(_))));

        let sparkline = Sparkline::new();
        let result = worksheet.add_sparkline_group(0, 5, 0, 5, &sparkline);
        assert!(matches!(result, Err(XlsxError::SparklineError(_))));

        let sparkline = Sparkline::new().set_range(("Sheet1", 0, 0, 0, 5));
        let result = worksheet.add_sparkline_group(0, 5, 0, 5, &sparkline);
        assert!(matches!(result, Err(XlsxError::SparklineError(_))));

        let sparkline = Sparkline::new().set_range(("Sheet1", 9, 9, 0, 5));
        let result = worksheet.add_sparkline_group(0, 5, 0, 5, &sparkline);
        assert!(matches!(result, Err(XlsxError::ChartError(_))));

        let sparkline = Sparkline::new().set_range(("Sheet1", 0, 0, 4, 5));
        let result = worksheet.add_sparkline_group(0, 5, 0, 5, &sparkline);
        assert!(matches!(result, Err(XlsxError::SparklineError(_))));
    }

    #[test]
    fn sparkline02_1() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, -2)?;
        worksheet.write(0, 1, 2)?;
        worksheet.write(0, 2, 3)?;
        worksheet.write(0, 3, -1)?;
        worksheet.write(0, 4, 0)?;

        let sparkline = Sparkline::new().set_range(("Sheet1", 0, 0, 0, 4));

        worksheet.add_sparkline(0, 5, &sparkline)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1:E1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData>
                <row r="1" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A1">
                    <v>-2</v>
                  </c>
                  <c r="B1">
                    <v>2</v>
                  </c>
                  <c r="C1">
                    <v>3</v>
                  </c>
                  <c r="D1">
                    <v>-1</v>
                  </c>
                  <c r="E1">
                    <v>0</v>
                  </c>
                </row>
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A1:E1</xm:f>
                          <xm:sqref>F1</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                  </x14:sparklineGroups>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn sparkline03() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, -2)?;
        worksheet.write(0, 1, 2)?;
        worksheet.write(0, 2, 3)?;
        worksheet.write(0, 3, -1)?;
        worksheet.write(0, 4, 0)?;
        worksheet.write(1, 0, -2)?;
        worksheet.write(1, 1, 2)?;
        worksheet.write(1, 2, 3)?;
        worksheet.write(1, 3, -1)?;
        worksheet.write(1, 4, 0)?;

        let sparkline1 = Sparkline::new().set_range(("Sheet1", 0, 0, 0, 4));
        let sparkline2 = Sparkline::new().set_range(("Sheet1", 1, 0, 1, 4));

        worksheet.add_sparkline(0, 5, &sparkline1)?;
        worksheet.add_sparkline(1, 5, &sparkline2)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1:E2"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData>
                <row r="1" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A1">
                    <v>-2</v>
                  </c>
                  <c r="B1">
                    <v>2</v>
                  </c>
                  <c r="C1">
                    <v>3</v>
                  </c>
                  <c r="D1">
                    <v>-1</v>
                  </c>
                  <c r="E1">
                    <v>0</v>
                  </c>
                </row>
                <row r="2" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A2">
                    <v>-2</v>
                  </c>
                  <c r="B2">
                    <v>2</v>
                  </c>
                  <c r="C2">
                    <v>3</v>
                  </c>
                  <c r="D2">
                    <v>-1</v>
                  </c>
                  <c r="E2">
                    <v>0</v>
                  </c>
                </row>
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A2:E2</xm:f>
                          <xm:sqref>F2</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A1:E1</xm:f>
                          <xm:sqref>F1</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                  </x14:sparklineGroups>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn sparkline04() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, -2)?;
        worksheet.write(0, 1, 2)?;
        worksheet.write(0, 2, 3)?;
        worksheet.write(0, 3, -1)?;
        worksheet.write(0, 4, 0)?;

        let sparkline = Sparkline::new()
            .set_range(("Sheet1", 0, 0, 0, 4))
            .set_type(SparklineType::Column);

        worksheet.add_sparkline(0, 5, &sparkline)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1:E1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData>
                <row r="1" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A1">
                    <v>-2</v>
                  </c>
                  <c r="B1">
                    <v>2</v>
                  </c>
                  <c r="C1">
                    <v>3</v>
                  </c>
                  <c r="D1">
                    <v>-1</v>
                  </c>
                  <c r="E1">
                    <v>0</v>
                  </c>
                </row>
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:sparklineGroup type="column" displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A1:E1</xm:f>
                          <xm:sqref>F1</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                  </x14:sparklineGroups>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn sparkline05() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        worksheet.write(0, 0, -2)?;
        worksheet.write(0, 1, 2)?;
        worksheet.write(0, 2, 3)?;
        worksheet.write(0, 3, -1)?;
        worksheet.write(0, 4, 0)?;

        let sparkline = Sparkline::new()
            .set_range(("Sheet1", 0, 0, 0, 4))
            .set_type(SparklineType::WinLose);

        worksheet.add_sparkline(0, 5, &sparkline)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1:E1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData>
                <row r="1" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A1">
                    <v>-2</v>
                  </c>
                  <c r="B1">
                    <v>2</v>
                  </c>
                  <c r="C1">
                    <v>3</v>
                  </c>
                  <c r="D1">
                    <v>-1</v>
                  </c>
                  <c r="E1">
                    <v>0</v>
                  </c>
                </row>
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:sparklineGroup type="stacked" displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A1:E1</xm:f>
                          <xm:sqref>F1</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                  </x14:sparklineGroups>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn sparkline06() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data = [-2, 2, 3, -1, 0];

        worksheet.write_row(0, 0, data)?;
        worksheet.write_row(1, 0, data)?;

        let sparkline = Sparkline::new().set_range(("Sheet1", 0, 0, 1, 4));

        worksheet.add_sparkline_group(0, 5, 1, 5, &sparkline)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1:E2"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData>
                <row r="1" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A1">
                    <v>-2</v>
                  </c>
                  <c r="B1">
                    <v>2</v>
                  </c>
                  <c r="C1">
                    <v>3</v>
                  </c>
                  <c r="D1">
                    <v>-1</v>
                  </c>
                  <c r="E1">
                    <v>0</v>
                  </c>
                </row>
                <row r="2" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A2">
                    <v>-2</v>
                  </c>
                  <c r="B2">
                    <v>2</v>
                  </c>
                  <c r="C2">
                    <v>3</v>
                  </c>
                  <c r="D2">
                    <v>-1</v>
                  </c>
                  <c r="E2">
                    <v>0</v>
                  </c>
                </row>
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A1:E1</xm:f>
                          <xm:sqref>F1</xm:sqref>
                        </x14:sparkline>
                        <x14:sparkline>
                          <xm:f>Sheet1!A2:E2</xm:f>
                          <xm:sqref>F2</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                  </x14:sparklineGroups>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn sparkline07() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data = [-2, 2, 3, -1, 0];

        worksheet.write_row(0, 0, data)?;
        worksheet.write_row(1, 0, data)?;
        worksheet.write_row(2, 0, data)?;
        worksheet.write_row(3, 0, data)?;
        worksheet.write_row(4, 0, data)?;
        worksheet.write_row(5, 0, data)?;
        worksheet.write_row(6, 0, data)?;

        let sparkline1 = Sparkline::new()
            .set_range(("Sheet1", 0, 0, 0, 4))
            .set_type(SparklineType::Column)
            .show_high_point(true);

        let sparkline2 = Sparkline::new()
            .set_range(("Sheet1", 1, 0, 1, 4))
            .set_type(SparklineType::Column)
            .show_low_point(true);

        let sparkline3 = Sparkline::new()
            .set_range(("Sheet1", 2, 0, 2, 4))
            .set_type(SparklineType::Column)
            .show_negative_points(true);

        let sparkline4 = Sparkline::new()
            .set_range(("Sheet1", 3, 0, 3, 4))
            .set_type(SparklineType::Column)
            .show_first_point(true);

        let sparkline5 = Sparkline::new()
            .set_range(("Sheet1", 4, 0, 4, 4))
            .set_type(SparklineType::Column)
            .show_last_point(true);

        let sparkline6 = Sparkline::new()
            .set_range(("Sheet1", 5, 0, 5, 4))
            .set_type(SparklineType::Column)
            .show_markers(true);

        let sparkline7 = Sparkline::new()
            .set_range(("Sheet1", 6, 0, 6, 4))
            .set_type(SparklineType::Column)
            .show_markers(true)
            .show_high_point(true)
            .show_low_point(true)
            .show_first_point(true)
            .show_last_point(true)
            .show_negative_points(true);

        worksheet.add_sparkline(0, 5, &sparkline1)?;
        worksheet.add_sparkline(1, 5, &sparkline2)?;
        worksheet.add_sparkline(2, 5, &sparkline3)?;
        worksheet.add_sparkline(3, 5, &sparkline4)?;
        worksheet.add_sparkline(4, 5, &sparkline5)?;
        worksheet.add_sparkline(5, 5, &sparkline6)?;
        worksheet.add_sparkline(6, 5, &sparkline7)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1:E7"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData>
                <row r="1" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A1">
                    <v>-2</v>
                  </c>
                  <c r="B1">
                    <v>2</v>
                  </c>
                  <c r="C1">
                    <v>3</v>
                  </c>
                  <c r="D1">
                    <v>-1</v>
                  </c>
                  <c r="E1">
                    <v>0</v>
                  </c>
                </row>
                <row r="2" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A2">
                    <v>-2</v>
                  </c>
                  <c r="B2">
                    <v>2</v>
                  </c>
                  <c r="C2">
                    <v>3</v>
                  </c>
                  <c r="D2">
                    <v>-1</v>
                  </c>
                  <c r="E2">
                    <v>0</v>
                  </c>
                </row>
                <row r="3" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A3">
                    <v>-2</v>
                  </c>
                  <c r="B3">
                    <v>2</v>
                  </c>
                  <c r="C3">
                    <v>3</v>
                  </c>
                  <c r="D3">
                    <v>-1</v>
                  </c>
                  <c r="E3">
                    <v>0</v>
                  </c>
                </row>
                <row r="4" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A4">
                    <v>-2</v>
                  </c>
                  <c r="B4">
                    <v>2</v>
                  </c>
                  <c r="C4">
                    <v>3</v>
                  </c>
                  <c r="D4">
                    <v>-1</v>
                  </c>
                  <c r="E4">
                    <v>0</v>
                  </c>
                </row>
                <row r="5" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A5">
                    <v>-2</v>
                  </c>
                  <c r="B5">
                    <v>2</v>
                  </c>
                  <c r="C5">
                    <v>3</v>
                  </c>
                  <c r="D5">
                    <v>-1</v>
                  </c>
                  <c r="E5">
                    <v>0</v>
                  </c>
                </row>
                <row r="6" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A6">
                    <v>-2</v>
                  </c>
                  <c r="B6">
                    <v>2</v>
                  </c>
                  <c r="C6">
                    <v>3</v>
                  </c>
                  <c r="D6">
                    <v>-1</v>
                  </c>
                  <c r="E6">
                    <v>0</v>
                  </c>
                </row>
                <row r="7" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A7">
                    <v>-2</v>
                  </c>
                  <c r="B7">
                    <v>2</v>
                  </c>
                  <c r="C7">
                    <v>3</v>
                  </c>
                  <c r="D7">
                    <v>-1</v>
                  </c>
                  <c r="E7">
                    <v>0</v>
                  </c>
                </row>
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" markers="1" high="1" low="1" first="1" last="1" negative="1">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A7:E7</xm:f>
                          <xm:sqref>F7</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" markers="1">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A6:E6</xm:f>
                          <xm:sqref>F6</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" last="1">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A5:E5</xm:f>
                          <xm:sqref>F5</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" first="1">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A4:E4</xm:f>
                          <xm:sqref>F4</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" negative="1">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A3:E3</xm:f>
                          <xm:sqref>F3</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" low="1">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A2:E2</xm:f>
                          <xm:sqref>F2</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" high="1">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A1:E1</xm:f>
                          <xm:sqref>F1</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                  </x14:sparklineGroups>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn sparkline08() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data = [-2, 2, 3, -1, 0];

        worksheet.write_row(0, 0, data)?;
        worksheet.write_row(1, 0, data)?;
        worksheet.write_row(2, 0, data)?;

        let sparkline1 = Sparkline::new()
            .set_range(("Sheet1", 0, 0, 0, 4))
            .set_style(1);

        let sparkline2 = Sparkline::new()
            .set_range(("Sheet1", 1, 0, 1, 4))
            .set_style(2);

        let sparkline3 = Sparkline::new()
            .set_range(("Sheet1", 2, 0, 2, 4))
            .set_sparkline_color(Color::RGB(0xFF0000));

        worksheet.add_sparkline(0, 5, &sparkline1)?;
        worksheet.add_sparkline(1, 5, &sparkline2)?;
        worksheet.add_sparkline(2, 5, &sparkline3)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1:E3"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData>
                <row r="1" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A1">
                    <v>-2</v>
                  </c>
                  <c r="B1">
                    <v>2</v>
                  </c>
                  <c r="C1">
                    <v>3</v>
                  </c>
                  <c r="D1">
                    <v>-1</v>
                  </c>
                  <c r="E1">
                    <v>0</v>
                  </c>
                </row>
                <row r="2" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A2">
                    <v>-2</v>
                  </c>
                  <c r="B2">
                    <v>2</v>
                  </c>
                  <c r="C2">
                    <v>3</v>
                  </c>
                  <c r="D2">
                    <v>-1</v>
                  </c>
                  <c r="E2">
                    <v>0</v>
                  </c>
                </row>
                <row r="3" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A3">
                    <v>-2</v>
                  </c>
                  <c r="B3">
                    <v>2</v>
                  </c>
                  <c r="C3">
                    <v>3</v>
                  </c>
                  <c r="D3">
                    <v>-1</v>
                  </c>
                  <c r="E3">
                    <v>0</v>
                  </c>
                </row>
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries rgb="FFFF0000"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A3:E3</xm:f>
                          <xm:sqref>F3</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="5" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="6"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="5" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="5" tint="0.39997558519241921"/>
                      <x14:colorLast theme="5" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="5"/>
                      <x14:colorLow theme="5"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A2:E2</xm:f>
                          <xm:sqref>F2</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A1:E1</xm:f>
                          <xm:sqref>F1</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                  </x14:sparklineGroups>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn sparkline09() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data = [-2, 2, 3, -1, 0];

        for row in 0..36 {
            worksheet.write_row(row, 0, data)?;

            let sparkline = Sparkline::new()
                .set_range(("Sheet1", row, 0, row, 4))
                .set_style((row + 1) as u8);

            worksheet.add_sparkline(row, 5, &sparkline)?;
        }

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1:E36"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData>
                <row r="1" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A1">
                    <v>-2</v>
                  </c>
                  <c r="B1">
                    <v>2</v>
                  </c>
                  <c r="C1">
                    <v>3</v>
                  </c>
                  <c r="D1">
                    <v>-1</v>
                  </c>
                  <c r="E1">
                    <v>0</v>
                  </c>
                </row>
                <row r="2" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A2">
                    <v>-2</v>
                  </c>
                  <c r="B2">
                    <v>2</v>
                  </c>
                  <c r="C2">
                    <v>3</v>
                  </c>
                  <c r="D2">
                    <v>-1</v>
                  </c>
                  <c r="E2">
                    <v>0</v>
                  </c>
                </row>
                <row r="3" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A3">
                    <v>-2</v>
                  </c>
                  <c r="B3">
                    <v>2</v>
                  </c>
                  <c r="C3">
                    <v>3</v>
                  </c>
                  <c r="D3">
                    <v>-1</v>
                  </c>
                  <c r="E3">
                    <v>0</v>
                  </c>
                </row>
                <row r="4" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A4">
                    <v>-2</v>
                  </c>
                  <c r="B4">
                    <v>2</v>
                  </c>
                  <c r="C4">
                    <v>3</v>
                  </c>
                  <c r="D4">
                    <v>-1</v>
                  </c>
                  <c r="E4">
                    <v>0</v>
                  </c>
                </row>
                <row r="5" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A5">
                    <v>-2</v>
                  </c>
                  <c r="B5">
                    <v>2</v>
                  </c>
                  <c r="C5">
                    <v>3</v>
                  </c>
                  <c r="D5">
                    <v>-1</v>
                  </c>
                  <c r="E5">
                    <v>0</v>
                  </c>
                </row>
                <row r="6" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A6">
                    <v>-2</v>
                  </c>
                  <c r="B6">
                    <v>2</v>
                  </c>
                  <c r="C6">
                    <v>3</v>
                  </c>
                  <c r="D6">
                    <v>-1</v>
                  </c>
                  <c r="E6">
                    <v>0</v>
                  </c>
                </row>
                <row r="7" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A7">
                    <v>-2</v>
                  </c>
                  <c r="B7">
                    <v>2</v>
                  </c>
                  <c r="C7">
                    <v>3</v>
                  </c>
                  <c r="D7">
                    <v>-1</v>
                  </c>
                  <c r="E7">
                    <v>0</v>
                  </c>
                </row>
                <row r="8" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A8">
                    <v>-2</v>
                  </c>
                  <c r="B8">
                    <v>2</v>
                  </c>
                  <c r="C8">
                    <v>3</v>
                  </c>
                  <c r="D8">
                    <v>-1</v>
                  </c>
                  <c r="E8">
                    <v>0</v>
                  </c>
                </row>
                <row r="9" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A9">
                    <v>-2</v>
                  </c>
                  <c r="B9">
                    <v>2</v>
                  </c>
                  <c r="C9">
                    <v>3</v>
                  </c>
                  <c r="D9">
                    <v>-1</v>
                  </c>
                  <c r="E9">
                    <v>0</v>
                  </c>
                </row>
                <row r="10" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A10">
                    <v>-2</v>
                  </c>
                  <c r="B10">
                    <v>2</v>
                  </c>
                  <c r="C10">
                    <v>3</v>
                  </c>
                  <c r="D10">
                    <v>-1</v>
                  </c>
                  <c r="E10">
                    <v>0</v>
                  </c>
                </row>
                <row r="11" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A11">
                    <v>-2</v>
                  </c>
                  <c r="B11">
                    <v>2</v>
                  </c>
                  <c r="C11">
                    <v>3</v>
                  </c>
                  <c r="D11">
                    <v>-1</v>
                  </c>
                  <c r="E11">
                    <v>0</v>
                  </c>
                </row>
                <row r="12" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A12">
                    <v>-2</v>
                  </c>
                  <c r="B12">
                    <v>2</v>
                  </c>
                  <c r="C12">
                    <v>3</v>
                  </c>
                  <c r="D12">
                    <v>-1</v>
                  </c>
                  <c r="E12">
                    <v>0</v>
                  </c>
                </row>
                <row r="13" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A13">
                    <v>-2</v>
                  </c>
                  <c r="B13">
                    <v>2</v>
                  </c>
                  <c r="C13">
                    <v>3</v>
                  </c>
                  <c r="D13">
                    <v>-1</v>
                  </c>
                  <c r="E13">
                    <v>0</v>
                  </c>
                </row>
                <row r="14" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A14">
                    <v>-2</v>
                  </c>
                  <c r="B14">
                    <v>2</v>
                  </c>
                  <c r="C14">
                    <v>3</v>
                  </c>
                  <c r="D14">
                    <v>-1</v>
                  </c>
                  <c r="E14">
                    <v>0</v>
                  </c>
                </row>
                <row r="15" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A15">
                    <v>-2</v>
                  </c>
                  <c r="B15">
                    <v>2</v>
                  </c>
                  <c r="C15">
                    <v>3</v>
                  </c>
                  <c r="D15">
                    <v>-1</v>
                  </c>
                  <c r="E15">
                    <v>0</v>
                  </c>
                </row>
                <row r="16" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A16">
                    <v>-2</v>
                  </c>
                  <c r="B16">
                    <v>2</v>
                  </c>
                  <c r="C16">
                    <v>3</v>
                  </c>
                  <c r="D16">
                    <v>-1</v>
                  </c>
                  <c r="E16">
                    <v>0</v>
                  </c>
                </row>
                <row r="17" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A17">
                    <v>-2</v>
                  </c>
                  <c r="B17">
                    <v>2</v>
                  </c>
                  <c r="C17">
                    <v>3</v>
                  </c>
                  <c r="D17">
                    <v>-1</v>
                  </c>
                  <c r="E17">
                    <v>0</v>
                  </c>
                </row>
                <row r="18" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A18">
                    <v>-2</v>
                  </c>
                  <c r="B18">
                    <v>2</v>
                  </c>
                  <c r="C18">
                    <v>3</v>
                  </c>
                  <c r="D18">
                    <v>-1</v>
                  </c>
                  <c r="E18">
                    <v>0</v>
                  </c>
                </row>
                <row r="19" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A19">
                    <v>-2</v>
                  </c>
                  <c r="B19">
                    <v>2</v>
                  </c>
                  <c r="C19">
                    <v>3</v>
                  </c>
                  <c r="D19">
                    <v>-1</v>
                  </c>
                  <c r="E19">
                    <v>0</v>
                  </c>
                </row>
                <row r="20" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A20">
                    <v>-2</v>
                  </c>
                  <c r="B20">
                    <v>2</v>
                  </c>
                  <c r="C20">
                    <v>3</v>
                  </c>
                  <c r="D20">
                    <v>-1</v>
                  </c>
                  <c r="E20">
                    <v>0</v>
                  </c>
                </row>
                <row r="21" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A21">
                    <v>-2</v>
                  </c>
                  <c r="B21">
                    <v>2</v>
                  </c>
                  <c r="C21">
                    <v>3</v>
                  </c>
                  <c r="D21">
                    <v>-1</v>
                  </c>
                  <c r="E21">
                    <v>0</v>
                  </c>
                </row>
                <row r="22" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A22">
                    <v>-2</v>
                  </c>
                  <c r="B22">
                    <v>2</v>
                  </c>
                  <c r="C22">
                    <v>3</v>
                  </c>
                  <c r="D22">
                    <v>-1</v>
                  </c>
                  <c r="E22">
                    <v>0</v>
                  </c>
                </row>
                <row r="23" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A23">
                    <v>-2</v>
                  </c>
                  <c r="B23">
                    <v>2</v>
                  </c>
                  <c r="C23">
                    <v>3</v>
                  </c>
                  <c r="D23">
                    <v>-1</v>
                  </c>
                  <c r="E23">
                    <v>0</v>
                  </c>
                </row>
                <row r="24" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A24">
                    <v>-2</v>
                  </c>
                  <c r="B24">
                    <v>2</v>
                  </c>
                  <c r="C24">
                    <v>3</v>
                  </c>
                  <c r="D24">
                    <v>-1</v>
                  </c>
                  <c r="E24">
                    <v>0</v>
                  </c>
                </row>
                <row r="25" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A25">
                    <v>-2</v>
                  </c>
                  <c r="B25">
                    <v>2</v>
                  </c>
                  <c r="C25">
                    <v>3</v>
                  </c>
                  <c r="D25">
                    <v>-1</v>
                  </c>
                  <c r="E25">
                    <v>0</v>
                  </c>
                </row>
                <row r="26" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A26">
                    <v>-2</v>
                  </c>
                  <c r="B26">
                    <v>2</v>
                  </c>
                  <c r="C26">
                    <v>3</v>
                  </c>
                  <c r="D26">
                    <v>-1</v>
                  </c>
                  <c r="E26">
                    <v>0</v>
                  </c>
                </row>
                <row r="27" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A27">
                    <v>-2</v>
                  </c>
                  <c r="B27">
                    <v>2</v>
                  </c>
                  <c r="C27">
                    <v>3</v>
                  </c>
                  <c r="D27">
                    <v>-1</v>
                  </c>
                  <c r="E27">
                    <v>0</v>
                  </c>
                </row>
                <row r="28" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A28">
                    <v>-2</v>
                  </c>
                  <c r="B28">
                    <v>2</v>
                  </c>
                  <c r="C28">
                    <v>3</v>
                  </c>
                  <c r="D28">
                    <v>-1</v>
                  </c>
                  <c r="E28">
                    <v>0</v>
                  </c>
                </row>
                <row r="29" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A29">
                    <v>-2</v>
                  </c>
                  <c r="B29">
                    <v>2</v>
                  </c>
                  <c r="C29">
                    <v>3</v>
                  </c>
                  <c r="D29">
                    <v>-1</v>
                  </c>
                  <c r="E29">
                    <v>0</v>
                  </c>
                </row>
                <row r="30" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A30">
                    <v>-2</v>
                  </c>
                  <c r="B30">
                    <v>2</v>
                  </c>
                  <c r="C30">
                    <v>3</v>
                  </c>
                  <c r="D30">
                    <v>-1</v>
                  </c>
                  <c r="E30">
                    <v>0</v>
                  </c>
                </row>
                <row r="31" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A31">
                    <v>-2</v>
                  </c>
                  <c r="B31">
                    <v>2</v>
                  </c>
                  <c r="C31">
                    <v>3</v>
                  </c>
                  <c r="D31">
                    <v>-1</v>
                  </c>
                  <c r="E31">
                    <v>0</v>
                  </c>
                </row>
                <row r="32" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A32">
                    <v>-2</v>
                  </c>
                  <c r="B32">
                    <v>2</v>
                  </c>
                  <c r="C32">
                    <v>3</v>
                  </c>
                  <c r="D32">
                    <v>-1</v>
                  </c>
                  <c r="E32">
                    <v>0</v>
                  </c>
                </row>
                <row r="33" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A33">
                    <v>-2</v>
                  </c>
                  <c r="B33">
                    <v>2</v>
                  </c>
                  <c r="C33">
                    <v>3</v>
                  </c>
                  <c r="D33">
                    <v>-1</v>
                  </c>
                  <c r="E33">
                    <v>0</v>
                  </c>
                </row>
                <row r="34" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A34">
                    <v>-2</v>
                  </c>
                  <c r="B34">
                    <v>2</v>
                  </c>
                  <c r="C34">
                    <v>3</v>
                  </c>
                  <c r="D34">
                    <v>-1</v>
                  </c>
                  <c r="E34">
                    <v>0</v>
                  </c>
                </row>
                <row r="35" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A35">
                    <v>-2</v>
                  </c>
                  <c r="B35">
                    <v>2</v>
                  </c>
                  <c r="C35">
                    <v>3</v>
                  </c>
                  <c r="D35">
                    <v>-1</v>
                  </c>
                  <c r="E35">
                    <v>0</v>
                  </c>
                </row>
                <row r="36" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A36">
                    <v>-2</v>
                  </c>
                  <c r="B36">
                    <v>2</v>
                  </c>
                  <c r="C36">
                    <v>3</v>
                  </c>
                  <c r="D36">
                    <v>-1</v>
                  </c>
                  <c r="E36">
                    <v>0</v>
                  </c>
                </row>
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="1"/>
                      <x14:colorNegative theme="9"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="8"/>
                      <x14:colorFirst theme="4"/>
                      <x14:colorLast theme="5"/>
                      <x14:colorHigh theme="6"/>
                      <x14:colorLow theme="7"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A36:E36</xm:f>
                          <xm:sqref>F36</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="3"/>
                      <x14:colorNegative theme="9"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="8"/>
                      <x14:colorFirst theme="4"/>
                      <x14:colorLast theme="5"/>
                      <x14:colorHigh theme="6"/>
                      <x14:colorLow theme="7"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A35:E35</xm:f>
                          <xm:sqref>F35</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries rgb="FF00B050"/>
                      <x14:colorNegative rgb="FFFF0000"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers rgb="FF0070C0"/>
                      <x14:colorFirst rgb="FFFFC000"/>
                      <x14:colorLast rgb="FFFFC000"/>
                      <x14:colorHigh rgb="FF00B050"/>
                      <x14:colorLow rgb="FFFF0000"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A34:E34</xm:f>
                          <xm:sqref>F34</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries rgb="FFC6EFCE"/>
                      <x14:colorNegative rgb="FFFFC7CE"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers rgb="FF8CADD6"/>
                      <x14:colorFirst rgb="FFFFDC47"/>
                      <x14:colorLast rgb="FFFFEB9C"/>
                      <x14:colorHigh rgb="FF60D276"/>
                      <x14:colorLow rgb="FFFF5367"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A33:E33</xm:f>
                          <xm:sqref>F33</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries rgb="FF5687C2"/>
                      <x14:colorNegative rgb="FFFFB620"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers rgb="FFD70077"/>
                      <x14:colorFirst rgb="FF777777"/>
                      <x14:colorLast rgb="FF359CEB"/>
                      <x14:colorHigh rgb="FF56BE79"/>
                      <x14:colorLow rgb="FFFF5055"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A32:E32</xm:f>
                          <xm:sqref>F32</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries rgb="FF5F5F5F"/>
                      <x14:colorNegative rgb="FFFFB620"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers rgb="FFD70077"/>
                      <x14:colorFirst rgb="FF5687C2"/>
                      <x14:colorLast rgb="FF359CEB"/>
                      <x14:colorHigh rgb="FF56BE79"/>
                      <x14:colorLow rgb="FFFF5055"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A31:E31</xm:f>
                          <xm:sqref>F31</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries rgb="FF0070C0"/>
                      <x14:colorNegative rgb="FF000000"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers rgb="FF000000"/>
                      <x14:colorFirst rgb="FF000000"/>
                      <x14:colorLast rgb="FF000000"/>
                      <x14:colorHigh rgb="FF000000"/>
                      <x14:colorLow rgb="FF000000"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A30:E30</xm:f>
                          <xm:sqref>F30</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries rgb="FF376092"/>
                      <x14:colorNegative rgb="FFD00000"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers rgb="FFD00000"/>
                      <x14:colorFirst rgb="FFD00000"/>
                      <x14:colorLast rgb="FFD00000"/>
                      <x14:colorHigh rgb="FFD00000"/>
                      <x14:colorLow rgb="FFD00000"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A29:E29</xm:f>
                          <xm:sqref>F29</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries rgb="FF000000"/>
                      <x14:colorNegative rgb="FF0070C0"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers rgb="FF0070C0"/>
                      <x14:colorFirst rgb="FF0070C0"/>
                      <x14:colorLast rgb="FF0070C0"/>
                      <x14:colorHigh rgb="FF0070C0"/>
                      <x14:colorLow rgb="FF0070C0"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A28:E28</xm:f>
                          <xm:sqref>F28</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries rgb="FF323232"/>
                      <x14:colorNegative rgb="FFD00000"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers rgb="FFD00000"/>
                      <x14:colorFirst rgb="FFD00000"/>
                      <x14:colorLast rgb="FFD00000"/>
                      <x14:colorHigh rgb="FFD00000"/>
                      <x14:colorLow rgb="FFD00000"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A27:E27</xm:f>
                          <xm:sqref>F27</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="1" tint="0.34998626667073579"/>
                      <x14:colorNegative theme="0" tint="-0.249977111117893"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="0" tint="-0.249977111117893"/>
                      <x14:colorFirst theme="0" tint="-0.249977111117893"/>
                      <x14:colorLast theme="0" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="0" tint="-0.249977111117893"/>
                      <x14:colorLow theme="0" tint="-0.249977111117893"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A26:E26</xm:f>
                          <xm:sqref>F26</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="1" tint="0.499984740745262"/>
                      <x14:colorNegative theme="1" tint="0.249977111117893"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="1" tint="0.249977111117893"/>
                      <x14:colorFirst theme="1" tint="0.249977111117893"/>
                      <x14:colorLast theme="1" tint="0.249977111117893"/>
                      <x14:colorHigh theme="1" tint="0.249977111117893"/>
                      <x14:colorLow theme="1" tint="0.249977111117893"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A25:E25</xm:f>
                          <xm:sqref>F25</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="9" tint="0.39997558519241921"/>
                      <x14:colorNegative theme="0" tint="-0.499984740745262"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="9" tint="0.79998168889431442"/>
                      <x14:colorFirst theme="9" tint="-0.249977111117893"/>
                      <x14:colorLast theme="9" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="9" tint="-0.499984740745262"/>
                      <x14:colorLow theme="9" tint="-0.499984740745262"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A24:E24</xm:f>
                          <xm:sqref>F24</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="8" tint="0.39997558519241921"/>
                      <x14:colorNegative theme="0" tint="-0.499984740745262"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="8" tint="0.79998168889431442"/>
                      <x14:colorFirst theme="8" tint="-0.249977111117893"/>
                      <x14:colorLast theme="8" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="8" tint="-0.499984740745262"/>
                      <x14:colorLow theme="8" tint="-0.499984740745262"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A23:E23</xm:f>
                          <xm:sqref>F23</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="7" tint="0.39997558519241921"/>
                      <x14:colorNegative theme="0" tint="-0.499984740745262"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="7" tint="0.79998168889431442"/>
                      <x14:colorFirst theme="7" tint="-0.249977111117893"/>
                      <x14:colorLast theme="7" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="7" tint="-0.499984740745262"/>
                      <x14:colorLow theme="7" tint="-0.499984740745262"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A22:E22</xm:f>
                          <xm:sqref>F22</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="6" tint="0.39997558519241921"/>
                      <x14:colorNegative theme="0" tint="-0.499984740745262"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="6" tint="0.79998168889431442"/>
                      <x14:colorFirst theme="6" tint="-0.249977111117893"/>
                      <x14:colorLast theme="6" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="6" tint="-0.499984740745262"/>
                      <x14:colorLow theme="6" tint="-0.499984740745262"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A21:E21</xm:f>
                          <xm:sqref>F21</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="5" tint="0.39997558519241921"/>
                      <x14:colorNegative theme="0" tint="-0.499984740745262"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="5" tint="0.79998168889431442"/>
                      <x14:colorFirst theme="5" tint="-0.249977111117893"/>
                      <x14:colorLast theme="5" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="5" tint="-0.499984740745262"/>
                      <x14:colorLow theme="5" tint="-0.499984740745262"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A20:E20</xm:f>
                          <xm:sqref>F20</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="4" tint="0.39997558519241921"/>
                      <x14:colorNegative theme="0" tint="-0.499984740745262"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="0.79998168889431442"/>
                      <x14:colorFirst theme="4" tint="-0.249977111117893"/>
                      <x14:colorLast theme="4" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="4" tint="-0.499984740745262"/>
                      <x14:colorLow theme="4" tint="-0.499984740745262"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A19:E19</xm:f>
                          <xm:sqref>F19</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="9"/>
                      <x14:colorNegative theme="4"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="9" tint="-0.249977111117893"/>
                      <x14:colorFirst theme="9" tint="-0.249977111117893"/>
                      <x14:colorLast theme="9" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="9" tint="-0.249977111117893"/>
                      <x14:colorLow theme="9" tint="-0.249977111117893"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A18:E18</xm:f>
                          <xm:sqref>F18</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="8"/>
                      <x14:colorNegative theme="9"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="8" tint="-0.249977111117893"/>
                      <x14:colorFirst theme="8" tint="-0.249977111117893"/>
                      <x14:colorLast theme="8" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="8" tint="-0.249977111117893"/>
                      <x14:colorLow theme="8" tint="-0.249977111117893"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A17:E17</xm:f>
                          <xm:sqref>F17</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="7"/>
                      <x14:colorNegative theme="8"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="7" tint="-0.249977111117893"/>
                      <x14:colorFirst theme="7" tint="-0.249977111117893"/>
                      <x14:colorLast theme="7" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="7" tint="-0.249977111117893"/>
                      <x14:colorLow theme="7" tint="-0.249977111117893"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A16:E16</xm:f>
                          <xm:sqref>F16</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="6"/>
                      <x14:colorNegative theme="7"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="6" tint="-0.249977111117893"/>
                      <x14:colorFirst theme="6" tint="-0.249977111117893"/>
                      <x14:colorLast theme="6" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="6" tint="-0.249977111117893"/>
                      <x14:colorLow theme="6" tint="-0.249977111117893"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A15:E15</xm:f>
                          <xm:sqref>F15</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="5"/>
                      <x14:colorNegative theme="6"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="5" tint="-0.249977111117893"/>
                      <x14:colorFirst theme="5" tint="-0.249977111117893"/>
                      <x14:colorLast theme="5" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="5" tint="-0.249977111117893"/>
                      <x14:colorLow theme="5" tint="-0.249977111117893"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A14:E14</xm:f>
                          <xm:sqref>F14</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="4"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.249977111117893"/>
                      <x14:colorFirst theme="4" tint="-0.249977111117893"/>
                      <x14:colorLast theme="4" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="4" tint="-0.249977111117893"/>
                      <x14:colorLow theme="4" tint="-0.249977111117893"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A13:E13</xm:f>
                          <xm:sqref>F13</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="9" tint="-0.249977111117893"/>
                      <x14:colorNegative theme="4"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.249977111117893"/>
                      <x14:colorFirst theme="4" tint="-0.249977111117893"/>
                      <x14:colorLast theme="4" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="4" tint="-0.249977111117893"/>
                      <x14:colorLow theme="4" tint="-0.249977111117893"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A12:E12</xm:f>
                          <xm:sqref>F12</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="8" tint="-0.249977111117893"/>
                      <x14:colorNegative theme="9"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="9" tint="-0.249977111117893"/>
                      <x14:colorFirst theme="9" tint="-0.249977111117893"/>
                      <x14:colorLast theme="9" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="9" tint="-0.249977111117893"/>
                      <x14:colorLow theme="9" tint="-0.249977111117893"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A11:E11</xm:f>
                          <xm:sqref>F11</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="7" tint="-0.249977111117893"/>
                      <x14:colorNegative theme="8"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="8" tint="-0.249977111117893"/>
                      <x14:colorFirst theme="8" tint="-0.249977111117893"/>
                      <x14:colorLast theme="8" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="8" tint="-0.249977111117893"/>
                      <x14:colorLow theme="8" tint="-0.249977111117893"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A10:E10</xm:f>
                          <xm:sqref>F10</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="6" tint="-0.249977111117893"/>
                      <x14:colorNegative theme="7"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="7" tint="-0.249977111117893"/>
                      <x14:colorFirst theme="7" tint="-0.249977111117893"/>
                      <x14:colorLast theme="7" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="7" tint="-0.249977111117893"/>
                      <x14:colorLow theme="7" tint="-0.249977111117893"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A9:E9</xm:f>
                          <xm:sqref>F9</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="5" tint="-0.249977111117893"/>
                      <x14:colorNegative theme="6"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="6" tint="-0.249977111117893"/>
                      <x14:colorFirst theme="6" tint="-0.249977111117893"/>
                      <x14:colorLast theme="6" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="6" tint="-0.249977111117893"/>
                      <x14:colorLow theme="6" tint="-0.249977111117893"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A8:E8</xm:f>
                          <xm:sqref>F8</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="4" tint="-0.249977111117893"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="5" tint="-0.249977111117893"/>
                      <x14:colorFirst theme="5" tint="-0.249977111117893"/>
                      <x14:colorLast theme="5" tint="-0.249977111117893"/>
                      <x14:colorHigh theme="5" tint="-0.249977111117893"/>
                      <x14:colorLow theme="5" tint="-0.249977111117893"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A7:E7</xm:f>
                          <xm:sqref>F7</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="9" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="4"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="9" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="9" tint="0.39997558519241921"/>
                      <x14:colorLast theme="9" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="9"/>
                      <x14:colorLow theme="9"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A6:E6</xm:f>
                          <xm:sqref>F6</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="8" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="9"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="8" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="8" tint="0.39997558519241921"/>
                      <x14:colorLast theme="8" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="8"/>
                      <x14:colorLow theme="8"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A5:E5</xm:f>
                          <xm:sqref>F5</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="7" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="8"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="7" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="7" tint="0.39997558519241921"/>
                      <x14:colorLast theme="7" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="7"/>
                      <x14:colorLow theme="7"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A4:E4</xm:f>
                          <xm:sqref>F4</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="6" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="7"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="6" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="6" tint="0.39997558519241921"/>
                      <x14:colorLast theme="6" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="6"/>
                      <x14:colorLow theme="6"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A3:E3</xm:f>
                          <xm:sqref>F3</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="5" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="6"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="5" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="5" tint="0.39997558519241921"/>
                      <x14:colorLast theme="5" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="5"/>
                      <x14:colorLow theme="5"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A2:E2</xm:f>
                          <xm:sqref>F2</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A1:E1</xm:f>
                          <xm:sqref>F1</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                  </x14:sparklineGroups>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn sparkline10_1() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data = [-2, 2, 3, -1, 0];
        worksheet.write_row(0, 0, data)?;

        let sparkline = Sparkline::new()
            .set_range(("Sheet1", 0, 0, 0, 4))
            .show_high_point(true)
            .set_high_point_color("#FFFF00")
            .show_low_point(true)
            .set_low_point_color("#92D050")
            .show_negative_points(true)
            .set_negative_points_color("#FF00000")
            .show_first_point(true)
            .set_first_point_color("#00B050")
            .show_last_point(true)
            .set_last_point_color("#00B0F0")
            .show_markers(true)
            .set_markers_color("#FFC000")
            .show_negative_points(true)
            .set_negative_points_color("#FF0000")
            .set_sparkline_color("#C00000");

        worksheet.add_sparkline(0, 5, &sparkline)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1:E1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData>
                <row r="1" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A1">
                    <v>-2</v>
                  </c>
                  <c r="B1">
                    <v>2</v>
                  </c>
                  <c r="C1">
                    <v>3</v>
                  </c>
                  <c r="D1">
                    <v>-1</v>
                  </c>
                  <c r="E1">
                    <v>0</v>
                  </c>
                </row>
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:sparklineGroup displayEmptyCellsAs="gap" markers="1" high="1" low="1" first="1" last="1" negative="1">
                      <x14:colorSeries rgb="FFC00000"/>
                      <x14:colorNegative rgb="FFFF0000"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers rgb="FFFFC000"/>
                      <x14:colorFirst rgb="FF00B050"/>
                      <x14:colorLast rgb="FF00B0F0"/>
                      <x14:colorHigh rgb="FFFFFF00"/>
                      <x14:colorLow rgb="FF92D050"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A1:E1</xm:f>
                          <xm:sqref>F1</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                  </x14:sparklineGroups>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn sparkline10_2() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data = [-2, 2, 3, -1, 0];
        worksheet.write_row(0, 0, data)?;

        let sparkline = Sparkline::new()
            .set_range(("Sheet1", 0, 0, 0, 4))
            .set_high_point_color("#FFFF00")
            .set_low_point_color("#92D050")
            .set_negative_points_color("#FF00000")
            .set_first_point_color("#00B050")
            .set_last_point_color("#00B0F0")
            .set_markers_color("#FFC000")
            .set_negative_points_color("#FF0000")
            .set_sparkline_color("#C00000");

        worksheet.add_sparkline(0, 5, &sparkline)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1:E1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData>
                <row r="1" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A1">
                    <v>-2</v>
                  </c>
                  <c r="B1">
                    <v>2</v>
                  </c>
                  <c r="C1">
                    <v>3</v>
                  </c>
                  <c r="D1">
                    <v>-1</v>
                  </c>
                  <c r="E1">
                    <v>0</v>
                  </c>
                </row>
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:sparklineGroup displayEmptyCellsAs="gap" markers="1" high="1" low="1" first="1" last="1" negative="1">
                      <x14:colorSeries rgb="FFC00000"/>
                      <x14:colorNegative rgb="FFFF0000"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers rgb="FFFFC000"/>
                      <x14:colorFirst rgb="FF00B050"/>
                      <x14:colorLast rgb="FF00B0F0"/>
                      <x14:colorHigh rgb="FFFFFF00"/>
                      <x14:colorLow rgb="FF92D050"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A1:E1</xm:f>
                          <xm:sqref>F1</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                  </x14:sparklineGroups>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn sparkline11() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data = [-2, 2, 3, -1, 0];

        worksheet.write_row(0, 0, data)?;
        worksheet.write_row(1, 0, data)?;
        worksheet.write_row(2, 0, data)?;
        worksheet.write_row(3, 0, [1, 2, 3, 4, 5])?;

        let sparkline1 = Sparkline::new()
            .set_range(("Sheet1", 0, 0, 0, 4))
            .show_high_point(true)
            .show_low_point(true)
            .show_negative_points(true)
            .show_first_point(true)
            .show_last_point(true)
            .show_markers(true)
            .show_axis(true)
            .set_right_to_left(true)
            .show_empty_cells_as(ChartEmptyCells::Zero)
            .set_line_weight(0.25)
            .set_custom_min(-0.5)
            .set_custom_max(0.5);

        let sparkline2 = Sparkline::new()
            .set_range(("Sheet1", 1, 0, 1, 4))
            .show_empty_cells_as(ChartEmptyCells::Connected)
            .set_line_weight(2.25)
            .set_group_max(true)
            .set_group_min(true);

        let sparkline3 = Sparkline::new()
            .set_range(("Sheet1", 2, 0, 2, 4))
            .set_line_weight(6)
            .show_hidden_data(true)
            .set_custom_min(0)
            .set_group_max(true)
            .set_date_range(("Sheet1", 3, 0, 3, 4));

        worksheet.add_sparkline(0, 5, &sparkline1)?;
        worksheet.add_sparkline(1, 5, &sparkline2)?;
        worksheet.add_sparkline(2, 5, &sparkline3)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1:E4"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData>
                <row r="1" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A1">
                    <v>-2</v>
                  </c>
                  <c r="B1">
                    <v>2</v>
                  </c>
                  <c r="C1">
                    <v>3</v>
                  </c>
                  <c r="D1">
                    <v>-1</v>
                  </c>
                  <c r="E1">
                    <v>0</v>
                  </c>
                </row>
                <row r="2" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A2">
                    <v>-2</v>
                  </c>
                  <c r="B2">
                    <v>2</v>
                  </c>
                  <c r="C2">
                    <v>3</v>
                  </c>
                  <c r="D2">
                    <v>-1</v>
                  </c>
                  <c r="E2">
                    <v>0</v>
                  </c>
                </row>
                <row r="3" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A3">
                    <v>-2</v>
                  </c>
                  <c r="B3">
                    <v>2</v>
                  </c>
                  <c r="C3">
                    <v>3</v>
                  </c>
                  <c r="D3">
                    <v>-1</v>
                  </c>
                  <c r="E3">
                    <v>0</v>
                  </c>
                </row>
                <row r="4" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A4">
                    <v>1</v>
                  </c>
                  <c r="B4">
                    <v>2</v>
                  </c>
                  <c r="C4">
                    <v>3</v>
                  </c>
                  <c r="D4">
                    <v>4</v>
                  </c>
                  <c r="E4">
                    <v>5</v>
                  </c>
                </row>
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:sparklineGroup manualMin="0" lineWeight="6" dateAxis="1" displayEmptyCellsAs="gap" displayHidden="1" minAxisType="custom" maxAxisType="group">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <xm:f>Sheet1!A4:E4</xm:f>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A3:E3</xm:f>
                          <xm:sqref>F3</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup lineWeight="2.25" displayEmptyCellsAs="span" minAxisType="group" maxAxisType="group">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A2:E2</xm:f>
                          <xm:sqref>F2</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                    <x14:sparklineGroup manualMax="0.5" manualMin="-0.5" lineWeight="0.25" markers="1" high="1" low="1" first="1" last="1" negative="1" displayXAxis="1" minAxisType="custom" maxAxisType="custom" rightToLeft="1">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A1:E1</xm:f>
                          <xm:sqref>F1</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                  </x14:sparklineGroups>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn sparkline12() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();
        worksheet.set_selected(true);

        let data = [-2, 2, 3, -1, 0];

        worksheet.write_row(0, 0, data)?;

        let sparkline1 = Sparkline::new()
            .set_range(("Sheet1", 0, 0, 0, 4))
            .set_custom_min(0)
            .set_custom_max(4);

        worksheet.add_sparkline(0, 5, &sparkline1)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1:E1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData>
                <row r="1" spans="1:5" x14ac:dyDescent="0.25">
                  <c r="A1">
                    <v>-2</v>
                  </c>
                  <c r="B1">
                    <v>2</v>
                  </c>
                  <c r="C1">
                    <v>3</v>
                  </c>
                  <c r="D1">
                    <v>-1</v>
                  </c>
                  <c r="E1">
                    <v>0</v>
                  </c>
                </row>
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:sparklineGroup manualMax="4" manualMin="0" displayEmptyCellsAs="gap" minAxisType="custom" maxAxisType="custom">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A1:E1</xm:f>
                          <xm:sqref>F1</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                  </x14:sparklineGroups>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn sparkline13() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();

        let data = [1, 2, 3, 4, 5, 6];

        worksheet.write_column(0, 0, data)?;
        worksheet.write_column(0, 1, data)?;
        worksheet.write_column(0, 2, data)?;
        worksheet.write_column(0, 3, data)?;
        worksheet.write_column(0, 4, data)?;
        worksheet.write_column(0, 5, data)?;

        let sparkline = Sparkline::new().set_range(("Sheet1", 0, 0, 5, 5));

        worksheet.add_sparkline_group(6, 0, 6, 5, &sparkline)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1:F6"/>
              <sheetViews>
                <sheetView workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData>
                <row r="1" spans="1:6" x14ac:dyDescent="0.25">
                  <c r="A1">
                    <v>1</v>
                  </c>
                  <c r="B1">
                    <v>1</v>
                  </c>
                  <c r="C1">
                    <v>1</v>
                  </c>
                  <c r="D1">
                    <v>1</v>
                  </c>
                  <c r="E1">
                    <v>1</v>
                  </c>
                  <c r="F1">
                    <v>1</v>
                  </c>
                </row>
                <row r="2" spans="1:6" x14ac:dyDescent="0.25">
                  <c r="A2">
                    <v>2</v>
                  </c>
                  <c r="B2">
                    <v>2</v>
                  </c>
                  <c r="C2">
                    <v>2</v>
                  </c>
                  <c r="D2">
                    <v>2</v>
                  </c>
                  <c r="E2">
                    <v>2</v>
                  </c>
                  <c r="F2">
                    <v>2</v>
                  </c>
                </row>
                <row r="3" spans="1:6" x14ac:dyDescent="0.25">
                  <c r="A3">
                    <v>3</v>
                  </c>
                  <c r="B3">
                    <v>3</v>
                  </c>
                  <c r="C3">
                    <v>3</v>
                  </c>
                  <c r="D3">
                    <v>3</v>
                  </c>
                  <c r="E3">
                    <v>3</v>
                  </c>
                  <c r="F3">
                    <v>3</v>
                  </c>
                </row>
                <row r="4" spans="1:6" x14ac:dyDescent="0.25">
                  <c r="A4">
                    <v>4</v>
                  </c>
                  <c r="B4">
                    <v>4</v>
                  </c>
                  <c r="C4">
                    <v>4</v>
                  </c>
                  <c r="D4">
                    <v>4</v>
                  </c>
                  <c r="E4">
                    <v>4</v>
                  </c>
                  <c r="F4">
                    <v>4</v>
                  </c>
                </row>
                <row r="5" spans="1:6" x14ac:dyDescent="0.25">
                  <c r="A5">
                    <v>5</v>
                  </c>
                  <c r="B5">
                    <v>5</v>
                  </c>
                  <c r="C5">
                    <v>5</v>
                  </c>
                  <c r="D5">
                    <v>5</v>
                  </c>
                  <c r="E5">
                    <v>5</v>
                  </c>
                  <c r="F5">
                    <v>5</v>
                  </c>
                </row>
                <row r="6" spans="1:6" x14ac:dyDescent="0.25">
                  <c r="A6">
                    <v>6</v>
                  </c>
                  <c r="B6">
                    <v>6</v>
                  </c>
                  <c r="C6">
                    <v>6</v>
                  </c>
                  <c r="D6">
                    <v>6</v>
                  </c>
                  <c r="E6">
                    <v>6</v>
                  </c>
                  <c r="F6">
                    <v>6</v>
                  </c>
                </row>
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A1:F1</xm:f>
                          <xm:sqref>A7</xm:sqref>
                        </x14:sparkline>
                        <x14:sparkline>
                          <xm:f>Sheet1!A2:F2</xm:f>
                          <xm:sqref>B7</xm:sqref>
                        </x14:sparkline>
                        <x14:sparkline>
                          <xm:f>Sheet1!A3:F3</xm:f>
                          <xm:sqref>C7</xm:sqref>
                        </x14:sparkline>
                        <x14:sparkline>
                          <xm:f>Sheet1!A4:F4</xm:f>
                          <xm:sqref>D7</xm:sqref>
                        </x14:sparkline>
                        <x14:sparkline>
                          <xm:f>Sheet1!A5:F5</xm:f>
                          <xm:sqref>E7</xm:sqref>
                        </x14:sparkline>
                        <x14:sparkline>
                          <xm:f>Sheet1!A6:F6</xm:f>
                          <xm:sqref>F7</xm:sqref>
                        </x14:sparkline>
                      </x14:sparklines>
                    </x14:sparklineGroup>
                  </x14:sparklineGroups>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }

    #[test]
    fn sparkline14() -> Result<(), XlsxError> {
        let mut worksheet = Worksheet::new();

        let data = [1, 2, 3, 4, 5, 6];

        worksheet.write_column(0, 0, data)?;
        worksheet.write_column(0, 1, data)?;
        worksheet.write_column(0, 2, data)?;
        worksheet.write_column(0, 3, data)?;
        worksheet.write_column(0, 4, data)?;
        worksheet.write_column(0, 5, data)?;

        let sparkline = Sparkline::new()
            .set_range(("Sheet1", 0, 0, 5, 5))
            .set_column_order(true);

        worksheet.add_sparkline_group(6, 0, 6, 5, &sparkline)?;

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
              <dimension ref="A1:F6"/>
              <sheetViews>
                <sheetView workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
              <sheetData>
                <row r="1" spans="1:6" x14ac:dyDescent="0.25">
                  <c r="A1">
                    <v>1</v>
                  </c>
                  <c r="B1">
                    <v>1</v>
                  </c>
                  <c r="C1">
                    <v>1</v>
                  </c>
                  <c r="D1">
                    <v>1</v>
                  </c>
                  <c r="E1">
                    <v>1</v>
                  </c>
                  <c r="F1">
                    <v>1</v>
                  </c>
                </row>
                <row r="2" spans="1:6" x14ac:dyDescent="0.25">
                  <c r="A2">
                    <v>2</v>
                  </c>
                  <c r="B2">
                    <v>2</v>
                  </c>
                  <c r="C2">
                    <v>2</v>
                  </c>
                  <c r="D2">
                    <v>2</v>
                  </c>
                  <c r="E2">
                    <v>2</v>
                  </c>
                  <c r="F2">
                    <v>2</v>
                  </c>
                </row>
                <row r="3" spans="1:6" x14ac:dyDescent="0.25">
                  <c r="A3">
                    <v>3</v>
                  </c>
                  <c r="B3">
                    <v>3</v>
                  </c>
                  <c r="C3">
                    <v>3</v>
                  </c>
                  <c r="D3">
                    <v>3</v>
                  </c>
                  <c r="E3">
                    <v>3</v>
                  </c>
                  <c r="F3">
                    <v>3</v>
                  </c>
                </row>
                <row r="4" spans="1:6" x14ac:dyDescent="0.25">
                  <c r="A4">
                    <v>4</v>
                  </c>
                  <c r="B4">
                    <v>4</v>
                  </c>
                  <c r="C4">
                    <v>4</v>
                  </c>
                  <c r="D4">
                    <v>4</v>
                  </c>
                  <c r="E4">
                    <v>4</v>
                  </c>
                  <c r="F4">
                    <v>4</v>
                  </c>
                </row>
                <row r="5" spans="1:6" x14ac:dyDescent="0.25">
                  <c r="A5">
                    <v>5</v>
                  </c>
                  <c r="B5">
                    <v>5</v>
                  </c>
                  <c r="C5">
                    <v>5</v>
                  </c>
                  <c r="D5">
                    <v>5</v>
                  </c>
                  <c r="E5">
                    <v>5</v>
                  </c>
                  <c r="F5">
                    <v>5</v>
                  </c>
                </row>
                <row r="6" spans="1:6" x14ac:dyDescent="0.25">
                  <c r="A6">
                    <v>6</v>
                  </c>
                  <c r="B6">
                    <v>6</v>
                  </c>
                  <c r="C6">
                    <v>6</v>
                  </c>
                  <c r="D6">
                    <v>6</v>
                  </c>
                  <c r="E6">
                    <v>6</v>
                  </c>
                  <c r="F6">
                    <v>6</v>
                  </c>
                </row>
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <extLst>
                <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:sparklineGroup displayEmptyCellsAs="gap">
                      <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                      <x14:colorNegative theme="5"/>
                      <x14:colorAxis rgb="FF000000"/>
                      <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                      <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                      <x14:colorLast theme="4" tint="0.39997558519241921"/>
                      <x14:colorHigh theme="4"/>
                      <x14:colorLow theme="4"/>
                      <x14:sparklines>
                        <x14:sparkline>
                          <xm:f>Sheet1!A1:A6</xm:f>
                          <xm:sqref>A7</xm:sqref>
                        </x14:sparkline>
                        <x14:sparkline>
                          <xm:f>Sheet1!B1:B6</xm:f>
                          <xm:sqref>B7</xm:sqref>
                        </x14:sparkline>
                        <x14:sparkline>
                          <xm:f>Sheet1!C1:C6</xm:f>
                          <xm:sqref>C7</xm:sqref>
                        </x14:sparkline>
                        <x14:sparkline>
                          <xm:f>Sheet1!D1:D6</xm:f>
                          <xm:sqref>D7</xm:sqref>
                        </x14:sparkline>
                        <x14:sparkline>
                          <xm:f>Sheet1!E1:E6</xm:f>
                          <xm:sqref>E7</xm:sqref>
                        </x14:sparkline>
                        <x14:sparkline>
                          <xm:f>Sheet1!F1:F6</xm:f>
                          <xm:sqref>F7</xm:sqref>
                        </x14:sparkline>
                    </x14:sparklines>
                    </x14:sparklineGroup>
                  </x14:sparklineGroups>
                </ext>
              </extLst>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);

        Ok(())
    }
}
