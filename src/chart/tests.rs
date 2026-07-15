// chart unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod chart_tests {

    use crate::chart::{Chart, ChartRange, ChartSeries, ChartType, XlsxError};
    use crate::test_functions::xml_to_vec;
    use crate::{xmlwriter, ChartRangeCacheDataType};
    use pretty_assertions::assert_eq;

    #[test]
    fn test_validation() {
        // Check for chart without series.
        let mut chart = Chart::new(ChartType::Scatter);
        let result = chart.validate();
        assert!(matches!(result, Err(XlsxError::ChartError(_))));

        // Check for chart with empty series.
        let mut chart = Chart::new(ChartType::Scatter);
        let series = ChartSeries::new();
        chart.push_series(&series);
        let result = chart.validate();
        assert!(matches!(result, Err(XlsxError::ChartError(_))));

        // Check for Scatter chart with empty categories.
        let mut chart = Chart::new(ChartType::Scatter);
        chart.add_series().set_values("Sheet1!$B$1:$B$3");
        let result = chart.validate();
        assert!(matches!(result, Err(XlsxError::ChartError(_))));

        // Check the value range for rows reversed.
        let mut chart = Chart::new(ChartType::Scatter);
        chart
            .add_series()
            .set_categories("Sheet1!$A$1:$A$3")
            .set_values("Sheet1!$B$3:$B$1");
        let result = chart.validate();
        assert!(matches!(result, Err(XlsxError::ChartError(_))));

        // Check the value range for rows reversed.
        let mut chart = Chart::new(ChartType::Scatter);
        chart
            .add_series()
            .set_categories("Sheet1!$A$1:$A$3")
            .set_values("Sheet1!$C$1:$B$3");
        let result = chart.validate();
        assert!(matches!(result, Err(XlsxError::ChartError(_))));

        // Check the value range for row out of range.
        let mut chart = Chart::new(ChartType::Scatter);
        chart
            .add_series()
            .set_categories("Sheet1!$A$1:$A$3")
            .set_values("Sheet1!$B$1:$B$1048577");
        let result = chart.validate();
        assert!(matches!(result, Err(XlsxError::ChartError(_))));

        // Check the value range for col out of range.
        let mut chart = Chart::new(ChartType::Scatter);
        chart
            .add_series()
            .set_categories("Sheet1!$A$1:$A$3")
            .set_values("Sheet1!$B$1:$XFE$10");
        let result = chart.validate();
        assert!(matches!(result, Err(XlsxError::ChartError(_))));

        // Check the category range for validation error.
        let mut chart = Chart::new(ChartType::Scatter);
        chart
            .add_series()
            .set_categories("Sheet1!$A$3:$A$1")
            .set_values("Sheet1!$B$1:$B$3");
        let result = chart.validate();
        assert!(matches!(result, Err(XlsxError::ChartError(_))));
    }

    #[test]
    fn test_assemble() {
        let mut series1 = ChartSeries::new();

        let mut range1 = ChartRange::new_from_range("Sheet1", 0, 0, 4, 0);
        range1.set_cache(&["1", "2", "3", "4", "5"], ChartRangeCacheDataType::Number);

        let mut range2 = ChartRange::new_from_range("Sheet1", 0, 1, 4, 1);
        range2.set_cache(&["2", "4", "6", "8", "10"], ChartRangeCacheDataType::Number);

        let mut range3 = ChartRange::new_from_string("Sheet1!$A$1:$A$5");
        range3.set_cache(&["1", "2", "3", "4", "5"], ChartRangeCacheDataType::Number);

        let mut range4 = ChartRange::new_from_string("Sheet1!$C$1:$C$5");
        range4.set_cache(
            &["3", "6", "9", "12", "15"],
            ChartRangeCacheDataType::Number,
        );

        series1.set_categories(&range1).set_values(&range2);

        let mut series2 = ChartSeries::new();
        series2.set_categories(&range3).set_values(&range4);

        let mut chart = Chart::new(ChartType::Bar);
        chart.push_series(&series1).push_series(&series2);

        chart.set_axis_ids(64052224, 64055552);

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                    <c:lang val="en-US"/>
                    <c:chart>
                        <c:plotArea>
                        <c:layout/>
                        <c:barChart>
                            <c:barDir val="bar"/>
                            <c:grouping val="clustered"/>
                            <c:ser>
                            <c:idx val="0"/>
                            <c:order val="0"/>
                            <c:cat>
                                <c:numRef>
                                <c:f>Sheet1!$A$1:$A$5</c:f>
                                <c:numCache>
                                    <c:formatCode>General</c:formatCode>
                                    <c:ptCount val="5"/>
                                    <c:pt idx="0">
                                    <c:v>1</c:v>
                                    </c:pt>
                                    <c:pt idx="1">
                                    <c:v>2</c:v>
                                    </c:pt>
                                    <c:pt idx="2">
                                    <c:v>3</c:v>
                                    </c:pt>
                                    <c:pt idx="3">
                                    <c:v>4</c:v>
                                    </c:pt>
                                    <c:pt idx="4">
                                    <c:v>5</c:v>
                                    </c:pt>
                                </c:numCache>
                                </c:numRef>
                            </c:cat>
                            <c:val>
                                <c:numRef>
                                <c:f>Sheet1!$B$1:$B$5</c:f>
                                <c:numCache>
                                    <c:formatCode>General</c:formatCode>
                                    <c:ptCount val="5"/>
                                    <c:pt idx="0">
                                    <c:v>2</c:v>
                                    </c:pt>
                                    <c:pt idx="1">
                                    <c:v>4</c:v>
                                    </c:pt>
                                    <c:pt idx="2">
                                    <c:v>6</c:v>
                                    </c:pt>
                                    <c:pt idx="3">
                                    <c:v>8</c:v>
                                    </c:pt>
                                    <c:pt idx="4">
                                    <c:v>10</c:v>
                                    </c:pt>
                                </c:numCache>
                                </c:numRef>
                            </c:val>
                            </c:ser>
                            <c:ser>
                            <c:idx val="1"/>
                            <c:order val="1"/>
                            <c:cat>
                                <c:numRef>
                                <c:f>Sheet1!$A$1:$A$5</c:f>
                                <c:numCache>
                                    <c:formatCode>General</c:formatCode>
                                    <c:ptCount val="5"/>
                                    <c:pt idx="0">
                                    <c:v>1</c:v>
                                    </c:pt>
                                    <c:pt idx="1">
                                    <c:v>2</c:v>
                                    </c:pt>
                                    <c:pt idx="2">
                                    <c:v>3</c:v>
                                    </c:pt>
                                    <c:pt idx="3">
                                    <c:v>4</c:v>
                                    </c:pt>
                                    <c:pt idx="4">
                                    <c:v>5</c:v>
                                    </c:pt>
                                </c:numCache>
                                </c:numRef>
                            </c:cat>
                            <c:val>
                                <c:numRef>
                                <c:f>Sheet1!$C$1:$C$5</c:f>
                                <c:numCache>
                                    <c:formatCode>General</c:formatCode>
                                    <c:ptCount val="5"/>
                                    <c:pt idx="0">
                                    <c:v>3</c:v>
                                    </c:pt>
                                    <c:pt idx="1">
                                    <c:v>6</c:v>
                                    </c:pt>
                                    <c:pt idx="2">
                                    <c:v>9</c:v>
                                    </c:pt>
                                    <c:pt idx="3">
                                    <c:v>12</c:v>
                                    </c:pt>
                                    <c:pt idx="4">
                                    <c:v>15</c:v>
                                    </c:pt>
                                </c:numCache>
                                </c:numRef>
                            </c:val>
                            </c:ser>
                            <c:axId val="64052224"/>
                            <c:axId val="64055552"/>
                        </c:barChart>
                        <c:catAx>
                            <c:axId val="64052224"/>
                            <c:scaling>
                            <c:orientation val="minMax"/>
                            </c:scaling>
                            <c:axPos val="l"/>
                            <c:numFmt formatCode="General" sourceLinked="1"/>
                            <c:tickLblPos val="nextTo"/>
                            <c:crossAx val="64055552"/>
                            <c:crosses val="autoZero"/>
                            <c:auto val="1"/>
                            <c:lblAlgn val="ctr"/>
                            <c:lblOffset val="100"/>
                        </c:catAx>
                        <c:valAx>
                            <c:axId val="64055552"/>
                            <c:scaling>
                            <c:orientation val="minMax"/>
                            </c:scaling>
                            <c:axPos val="b"/>
                            <c:majorGridlines/>
                            <c:numFmt formatCode="General" sourceLinked="1"/>
                            <c:tickLblPos val="nextTo"/>
                            <c:crossAx val="64052224"/>
                            <c:crosses val="autoZero"/>
                            <c:crossBetween val="between"/>
                        </c:valAx>
                        </c:plotArea>
                        <c:legend>
                        <c:legendPos val="r"/>
                        <c:layout/>
                        </c:legend>
                        <c:plotVisOnly val="1"/>
                    </c:chart>
                    <c:printSettings>
                        <c:headerFooter/>
                        <c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>
                        <c:pageSetup/>
                    </c:printSettings>
                    </c:chartSpace>

                "#,
        );

        assert_eq!(expected, got);
    }

    // Create a chart of the given 3D type with a single data series and
    // fixed axis ids, and return the generated chart.xml output.
    fn assemble_3d_chart(chart_type: ChartType) -> Vec<String> {
        let mut range1 = ChartRange::new_from_range("Sheet1", 0, 0, 4, 0);
        range1.set_cache(&["1", "2", "3", "4", "5"], ChartRangeCacheDataType::Number);

        let mut range2 = ChartRange::new_from_range("Sheet1", 0, 1, 4, 1);
        range2.set_cache(&["2", "4", "6", "8", "10"], ChartRangeCacheDataType::Number);

        let mut series1 = ChartSeries::new();
        series1.set_categories(&range1).set_values(&range2);

        let mut chart = Chart::new(chart_type);
        chart.push_series(&series1);

        chart.set_axis_ids(64052224, 64055552);
        chart.set_axis3_id(64058880);

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);
        xml_to_vec(got)
    }

    // The single series shared by the 3D chart assembly tests.
    const CHART_3D_TEST_SERIES: &str = r#"
            <c:ser>
              <c:idx val="0"/>
              <c:order val="0"/>
              <c:cat>
                <c:numRef>
                  <c:f>Sheet1!$A$1:$A$5</c:f>
                  <c:numCache>
                    <c:formatCode>General</c:formatCode>
                    <c:ptCount val="5"/>
                    <c:pt idx="0">
                      <c:v>1</c:v>
                    </c:pt>
                    <c:pt idx="1">
                      <c:v>2</c:v>
                    </c:pt>
                    <c:pt idx="2">
                      <c:v>3</c:v>
                    </c:pt>
                    <c:pt idx="3">
                      <c:v>4</c:v>
                    </c:pt>
                    <c:pt idx="4">
                      <c:v>5</c:v>
                    </c:pt>
                  </c:numCache>
                </c:numRef>
              </c:cat>
              <c:val>
                <c:numRef>
                  <c:f>Sheet1!$B$1:$B$5</c:f>
                  <c:numCache>
                    <c:formatCode>General</c:formatCode>
                    <c:ptCount val="5"/>
                    <c:pt idx="0">
                      <c:v>2</c:v>
                    </c:pt>
                    <c:pt idx="1">
                      <c:v>4</c:v>
                    </c:pt>
                    <c:pt idx="2">
                      <c:v>6</c:v>
                    </c:pt>
                    <c:pt idx="3">
                      <c:v>8</c:v>
                    </c:pt>
                    <c:pt idx="4">
                      <c:v>10</c:v>
                    </c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
    "#;

    const CHART_3D_TEST_CAT_VAL_AXES: &str = r#"
            <c:catAx>
              <c:axId val="64052224"/>
              <c:scaling>
                <c:orientation val="minMax"/>
              </c:scaling>
              <c:axPos val="b"/>
              <c:numFmt formatCode="General" sourceLinked="1"/>
              <c:tickLblPos val="nextTo"/>
              <c:crossAx val="64055552"/>
              <c:crosses val="autoZero"/>
              <c:auto val="1"/>
              <c:lblAlgn val="ctr"/>
              <c:lblOffset val="100"/>
            </c:catAx>
            <c:valAx>
              <c:axId val="64055552"/>
              <c:scaling>
                <c:orientation val="minMax"/>
              </c:scaling>
              <c:axPos val="l"/>
              <c:majorGridlines/>
              <c:numFmt formatCode="General" sourceLinked="1"/>
              <c:tickLblPos val="nextTo"/>
              <c:crossAx val="64052224"/>
              <c:crosses val="autoZero"/>
              <c:crossBetween val="between"/>
            </c:valAx>
    "#;

    const CHART_3D_TEST_SER_AXIS: &str = r#"
            <c:serAx>
              <c:axId val="64058880"/>
              <c:scaling>
                <c:orientation val="minMax"/>
              </c:scaling>
              <c:axPos val="b"/>
              <c:tickLblPos val="nextTo"/>
              <c:crossAx val="64055552"/>
            </c:serAx>
    "#;

    const CHART_3D_TEST_FOOTER: &str = r#"
          </c:plotArea>
          <c:legend>
            <c:legendPos val="r"/>
            <c:layout/>
          </c:legend>
          <c:plotVisOnly val="1"/>
        </c:chart>
        <c:printSettings>
          <c:headerFooter/>
          <c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>
          <c:pageSetup/>
        </c:printSettings>
        </c:chartSpace>
    "#;

    const CHART_3D_TEST_HEADER: &str = r#"
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <c:lang val="en-US"/>
        <c:chart>
    "#;

    #[test]
    fn test_assemble_column_3d() {
        let got = assemble_3d_chart(ChartType::Column3D);

        let expected = xml_to_vec(&format!(
            r#"
            {header}
              <c:view3D>
                <c:rotX val="15"/>
                <c:rotY val="20"/>
                <c:rAngAx val="1"/>
              </c:view3D>
              <c:plotArea>
              <c:layout/>
              <c:bar3DChart>
                <c:barDir val="col"/>
                <c:grouping val="clustered"/>
                {series}
                <c:shape val="box"/>
                <c:axId val="64052224"/>
                <c:axId val="64055552"/>
                <c:axId val="0"/>
              </c:bar3DChart>
              {cat_val_axes}
            {footer}
            "#,
            header = CHART_3D_TEST_HEADER.trim(),
            series = CHART_3D_TEST_SERIES.trim(),
            cat_val_axes = CHART_3D_TEST_CAT_VAL_AXES.trim(),
            footer = CHART_3D_TEST_FOOTER.trim(),
        ));

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble_pie_3d() {
        let got = assemble_3d_chart(ChartType::Pie3D);

        let expected = xml_to_vec(&format!(
            r#"
            {header}
              <c:view3D>
                <c:rotX val="30"/>
                <c:rotY val="0"/>
                <c:rAngAx val="0"/>
                <c:perspective val="30"/>
              </c:view3D>
              <c:plotArea>
              <c:layout/>
              <c:pie3DChart>
                <c:varyColors val="1"/>
                {series}
              </c:pie3DChart>
            {footer}
            "#,
            header = CHART_3D_TEST_HEADER.trim(),
            series = CHART_3D_TEST_SERIES.trim(),
            footer = CHART_3D_TEST_FOOTER.trim(),
        ));

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble_line_3d() {
        let got = assemble_3d_chart(ChartType::Line3D);

        let expected = xml_to_vec(&format!(
            r#"
            {header}
              <c:view3D>
                <c:rotX val="15"/>
                <c:rotY val="20"/>
                <c:rAngAx val="1"/>
              </c:view3D>
              <c:plotArea>
              <c:layout/>
              <c:line3DChart>
                <c:grouping val="standard"/>
                <c:ser>
                  <c:idx val="0"/>
                  <c:order val="0"/>
                  <c:marker>
                    <c:symbol val="none"/>
                  </c:marker>
                  <c:cat>
                    <c:numRef>
                      <c:f>Sheet1!$A$1:$A$5</c:f>
                      <c:numCache>
                        <c:formatCode>General</c:formatCode>
                        <c:ptCount val="5"/>
                        <c:pt idx="0">
                          <c:v>1</c:v>
                        </c:pt>
                        <c:pt idx="1">
                          <c:v>2</c:v>
                        </c:pt>
                        <c:pt idx="2">
                          <c:v>3</c:v>
                        </c:pt>
                        <c:pt idx="3">
                          <c:v>4</c:v>
                        </c:pt>
                        <c:pt idx="4">
                          <c:v>5</c:v>
                        </c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:cat>
                  <c:val>
                    <c:numRef>
                      <c:f>Sheet1!$B$1:$B$5</c:f>
                      <c:numCache>
                        <c:formatCode>General</c:formatCode>
                        <c:ptCount val="5"/>
                        <c:pt idx="0">
                          <c:v>2</c:v>
                        </c:pt>
                        <c:pt idx="1">
                          <c:v>4</c:v>
                        </c:pt>
                        <c:pt idx="2">
                          <c:v>6</c:v>
                        </c:pt>
                        <c:pt idx="3">
                          <c:v>8</c:v>
                        </c:pt>
                        <c:pt idx="4">
                          <c:v>10</c:v>
                        </c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
                <c:axId val="64052224"/>
                <c:axId val="64055552"/>
                <c:axId val="64058880"/>
              </c:line3DChart>
              {cat_val_axes}
              {ser_axis}
            {footer}
            "#,
            header = CHART_3D_TEST_HEADER.trim(),
            cat_val_axes = CHART_3D_TEST_CAT_VAL_AXES.trim(),
            ser_axis = CHART_3D_TEST_SER_AXIS.trim(),
            footer = CHART_3D_TEST_FOOTER.trim(),
        ));

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble_surface_3d() {
        let got = assemble_3d_chart(ChartType::Surface3D);

        let expected = xml_to_vec(&format!(
            r#"
            {header}
              <c:view3D>
                <c:rotX val="15"/>
                <c:rotY val="20"/>
                <c:rAngAx val="0"/>
                <c:perspective val="30"/>
              </c:view3D>
              <c:plotArea>
              <c:layout/>
              <c:surface3DChart>
                <c:wireframe val="0"/>
                {series}
                <c:axId val="64052224"/>
                <c:axId val="64055552"/>
                <c:axId val="64058880"/>
              </c:surface3DChart>
              {cat_val_axes}
              {ser_axis}
            {footer}
            "#,
            header = CHART_3D_TEST_HEADER.trim(),
            series = CHART_3D_TEST_SERIES.trim(),
            cat_val_axes = CHART_3D_TEST_CAT_VAL_AXES.trim(),
            ser_axis = CHART_3D_TEST_SER_AXIS.trim(),
            footer = CHART_3D_TEST_FOOTER.trim(),
        ));

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble_contour() {
        let got = assemble_3d_chart(ChartType::Contour);

        let expected = xml_to_vec(&format!(
            r#"
            {header}
              <c:view3D>
                <c:rotX val="90"/>
                <c:rotY val="0"/>
                <c:rAngAx val="0"/>
                <c:perspective val="0"/>
              </c:view3D>
              <c:plotArea>
              <c:layout/>
              <c:surfaceChart>
                <c:wireframe val="0"/>
                {series}
                <c:axId val="64052224"/>
                <c:axId val="64055552"/>
                <c:axId val="64058880"/>
              </c:surfaceChart>
              {cat_val_axes}
              {ser_axis}
            {footer}
            "#,
            header = CHART_3D_TEST_HEADER.trim(),
            series = CHART_3D_TEST_SERIES.trim(),
            cat_val_axes = CHART_3D_TEST_CAT_VAL_AXES.trim(),
            ser_axis = CHART_3D_TEST_SER_AXIS.trim(),
            footer = CHART_3D_TEST_FOOTER.trim(),
        ));

        assert_eq!(expected, got);
    }

    #[test]
    fn test_view_3d_options() {
        let mut range1 = ChartRange::new_from_range("Sheet1", 0, 0, 4, 0);
        range1.set_cache(&["1", "2", "3", "4", "5"], ChartRangeCacheDataType::Number);

        let mut range2 = ChartRange::new_from_range("Sheet1", 0, 1, 4, 1);
        range2.set_cache(&["2", "4", "6", "8", "10"], ChartRangeCacheDataType::Number);

        let mut series1 = ChartSeries::new();
        series1.set_categories(&range1).set_values(&range2);

        let mut chart = Chart::new(ChartType::Column3D);
        chart.push_series(&series1);

        chart.set_axis_ids(64052224, 64055552);

        chart
            .set_view_3d_x_rotation(45)
            .set_view_3d_y_rotation(90)
            .set_view_3d_perspective(20)
            .set_view_3d_depth_percent(200)
            .set_view_3d_height_percent(150)
            .set_view_3d_right_angle_axes(false)
            .set_gap_depth(50);

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);

        // Check the c:view3D element reflects the configured view.
        assert!(got.contains(
            "<c:view3D>\
            <c:rotX val=\"45\"/>\
            <c:hPercent val=\"150\"/>\
            <c:rotY val=\"90\"/>\
            <c:depthPercent val=\"200\"/>\
            <c:rAngAx val=\"0\"/>\
            <c:perspective val=\"40\"/>\
            </c:view3D>"
        ));

        // Check the c:gapDepth element is written for non-default values.
        assert!(got.contains("<c:gapDepth val=\"50\"/>"));
    }

    #[test]
    fn test_range_from_string() {
        let range_string = "=Sheet1!$A$1:$A$5";
        let range = ChartRange::new_from_string(range_string);
        assert_eq!("Sheet1!$A$1:$A$5", range.formula_string());
        assert_eq!("Sheet1", range.sheet_name);

        let range_string = "Sheet1!$A$1:$A$5";
        let range = ChartRange::new_from_string(range_string);
        assert_eq!("Sheet1!$A$1:$A$5", range.formula_string());
        assert_eq!("Sheet1", range.sheet_name);

        let range_string = "Sheet 1!$A$1:$A$5";
        let range = ChartRange::new_from_string(range_string);
        assert_eq!("'Sheet 1'!$A$1:$A$5", range.formula_string());
        assert_eq!("Sheet 1", range.sheet_name);
    }
}
