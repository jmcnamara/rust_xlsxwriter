// chart unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod chart_tests {

    use crate::chart::{Chart, ChartRange, ChartRangeCacheData, ChartSeries, ChartType, XlsxError};
    use crate::test_functions::xml_to_vec;
    use crate::{
        xmlwriter, ChartDataLabel, ChartFormat, ChartMapProjection, ChartParentLabelLayout,
        ChartQuartileMethod, ChartRangeCacheDataType, ChartRegionLabelLayout, ChartSolidFill,
    };
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

    #[test]
    fn test_bubble_chart() {
        let mut x_values = ChartRange::new_from_range("Sheet1", 0, 0, 2, 0);
        x_values.set_cache(&["1", "2", "3"], ChartRangeCacheDataType::Number);

        let mut y_values = ChartRange::new_from_range("Sheet1", 0, 1, 2, 1);
        y_values.set_cache(&["10", "40", "30"], ChartRangeCacheDataType::Number);

        let mut sizes = ChartRange::new_from_range("Sheet1", 0, 2, 2, 2);
        sizes.set_cache(&["5", "12", "8"], ChartRangeCacheDataType::Number);

        let mut chart = Chart::new(ChartType::Bubble);
        chart
            .add_series()
            .set_categories(&x_values)
            .set_values(&y_values)
            .set_bubble_sizes(&sizes);

        chart.set_axis_ids(50010001, 50010002);

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);

        // The series contains X values, Y values and the bubble sizes.
        assert!(got.contains(
            "<c:bubbleChart>\
             <c:varyColors val=\"0\"/>\
             <c:ser>\
             <c:idx val=\"0\"/>\
             <c:order val=\"0\"/>\
             <c:xVal>"
        ));
        assert!(got.contains(
            "<c:bubbleSize>\
             <c:numRef>\
             <c:f>Sheet1!$C$1:$C$3</c:f>"
        ));
        assert!(got.contains(
            "</c:bubbleSize>\
             <c:bubble3D val=\"0\"/>\
             </c:ser>\
             <c:bubbleScale val=\"100\"/>\
             <c:showNegBubbles val=\"0\"/>\
             <c:axId val=\"50010001\"/>\
             <c:axId val=\"50010002\"/>\
             </c:bubbleChart>"
        ));

        // Bubble charts have two value axes, like Scatter charts.
        assert_eq!(2, got.matches("<c:valAx>").count());
        assert!(!got.contains("<c:catAx>"));
    }

    #[test]
    fn test_bubble_chart_without_sizes() {
        // The c:bubbleSize element should be omitted if no sizes were set.
        let mut chart = Chart::new(ChartType::Bubble);
        chart
            .add_series()
            .set_categories("Sheet1!$A$1:$A$3")
            .set_values("Sheet1!$B$1:$B$3");

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);

        assert!(!got.contains("<c:bubbleSize>"));
        assert!(got.contains("<c:bubble3D val=\"0\"/>"));
    }

    #[test]
    fn test_of_pie_chart() {
        let mut categories = ChartRange::new_from_range("Sheet1", 0, 0, 2, 0);
        categories.set_cache(
            &["Apple", "Cherry", "Pecan"],
            ChartRangeCacheDataType::String,
        );

        let mut values = ChartRange::new_from_range("Sheet1", 0, 1, 2, 1);
        values.set_cache(&["60", "30", "10"], ChartRangeCacheDataType::Number);

        let mut chart = Chart::new(ChartType::PieOfPie);
        chart
            .add_series()
            .set_categories(&categories)
            .set_values(&values);

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);

        assert!(got.contains(
            "<c:ofPieChart>\
             <c:ofPieType val=\"pie\"/>\
             <c:varyColors val=\"1\"/>\
             <c:ser>"
        ));
        assert!(got.contains(
            "<c:gapWidth val=\"100\"/>\
             <c:serLines/>\
             </c:ofPieChart>"
        ));

        // Pie style charts don't have axes.
        assert!(!got.contains("<c:catAx>"));
        assert!(!got.contains("<c:valAx>"));

        // Check the Bar-of-Pie variant type.
        let mut chart = Chart::new(ChartType::BarOfPie);
        chart
            .add_series()
            .set_categories(&categories)
            .set_values(&values);

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);
        assert!(got.contains("<c:ofPieType val=\"bar\"/>"));
    }

    // -----------------------------------------------------------------------
    // ChartEx (Excel 2016+) chart tests.
    // -----------------------------------------------------------------------

    #[test]
    fn test_chartex_waterfall() {
        let mut categories = ChartRange::new_from_range("Sheet1", 0, 0, 5, 0);
        categories.set_cache(
            &["Start", "Q1", "Q2", "Q3", "Q4", "End"],
            ChartRangeCacheDataType::String,
        );

        let mut values = ChartRange::new_from_range("Sheet1", 0, 1, 5, 1);
        values.set_cache(
            &["100", "30", "-20", "40", "25", "175"],
            ChartRangeCacheDataType::Number,
        );

        let mut chart = Chart::new(ChartType::Waterfall);
        chart
            .add_series()
            .set_categories(&categories)
            .set_values(&values);

        chart
            .set_waterfall_subtotals(&[0, 5])
            .set_waterfall_connector_lines(false);

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <cx:chartData>
                <cx:data id="0">
                <cx:strDim type="cat">
                <cx:f>Sheet1!$A$1:$A$6</cx:f>
                <cx:lvl ptCount="6">
                <cx:pt idx="0">Start</cx:pt>
                <cx:pt idx="1">Q1</cx:pt>
                <cx:pt idx="2">Q2</cx:pt>
                <cx:pt idx="3">Q3</cx:pt>
                <cx:pt idx="4">Q4</cx:pt>
                <cx:pt idx="5">End</cx:pt>
                </cx:lvl>
                </cx:strDim>
                <cx:numDim type="val">
                <cx:f>Sheet1!$B$1:$B$6</cx:f>
                <cx:lvl ptCount="6" formatCode="General">
                <cx:pt idx="0">100</cx:pt>
                <cx:pt idx="1">30</cx:pt>
                <cx:pt idx="2">-20</cx:pt>
                <cx:pt idx="3">40</cx:pt>
                <cx:pt idx="4">25</cx:pt>
                <cx:pt idx="5">175</cx:pt>
                </cx:lvl>
                </cx:numDim>
                </cx:data>
                </cx:chartData>
                <cx:chart>
                <cx:plotArea>
                <cx:plotAreaRegion>
                <cx:series layoutId="waterfall" uniqueId="{00000000-0000-0000-0000-000000000001}" formatIdx="0">
                <cx:dataId val="0"/>
                <cx:layoutPr>
                <cx:visibility connectorLines="0"/>
                <cx:subtotals>
                <cx:idx val="0"/>
                <cx:idx val="5"/>
                </cx:subtotals>
                </cx:layoutPr>
                </cx:series>
                </cx:plotAreaRegion>
                <cx:axis id="0">
                <cx:catScaling gapWidth="0.5"/>
                <cx:tickLabels/>
                </cx:axis>
                <cx:axis id="1">
                <cx:valScaling/>
                <cx:majorGridlines/>
                <cx:tickLabels/>
                </cx:axis>
                </cx:plotArea>
                <cx:legend pos="r" align="ctr" overlay="0"/>
                </cx:chart>
                </cx:chartSpace>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_chartex_funnel() {
        let mut categories = ChartRange::new_from_range("Sheet1", 0, 0, 2, 0);
        categories.set_cache(
            &["Leads", "Meetings", "Sales"],
            ChartRangeCacheDataType::String,
        );

        let mut values = ChartRange::new_from_range("Sheet1", 0, 1, 2, 1);
        values.set_cache(&["1000", "400", "100"], ChartRangeCacheDataType::Number);

        let mut chart = Chart::new(ChartType::Funnel);
        chart
            .add_series()
            .set_categories(&categories)
            .set_values(&values)
            .set_name("Pipeline")
            .set_data_label(ChartDataLabel::new().show_value());

        chart.title().set_name("Funnel");
        chart.legend().set_hidden();

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <cx:chartData>
                <cx:data id="0">
                <cx:strDim type="cat">
                <cx:f>Sheet1!$A$1:$A$3</cx:f>
                <cx:lvl ptCount="3">
                <cx:pt idx="0">Leads</cx:pt>
                <cx:pt idx="1">Meetings</cx:pt>
                <cx:pt idx="2">Sales</cx:pt>
                </cx:lvl>
                </cx:strDim>
                <cx:numDim type="val">
                <cx:f>Sheet1!$B$1:$B$3</cx:f>
                <cx:lvl ptCount="3" formatCode="General">
                <cx:pt idx="0">1000</cx:pt>
                <cx:pt idx="1">400</cx:pt>
                <cx:pt idx="2">100</cx:pt>
                </cx:lvl>
                </cx:numDim>
                </cx:data>
                </cx:chartData>
                <cx:chart>
                <cx:title pos="t" align="ctr" overlay="0">
                <cx:tx>
                <cx:txData>
                <cx:v>Funnel</cx:v>
                </cx:txData>
                </cx:tx>
                </cx:title>
                <cx:plotArea>
                <cx:plotAreaRegion>
                <cx:series layoutId="funnel" uniqueId="{00000000-0000-0000-0000-000000000001}" formatIdx="0">
                <cx:tx>
                <cx:txData>
                <cx:v>Pipeline</cx:v>
                </cx:txData>
                </cx:tx>
                <cx:dataLabels>
                <cx:visibility seriesName="0" categoryName="0" value="1"/>
                </cx:dataLabels>
                <cx:dataId val="0"/>
                </cx:series>
                </cx:plotAreaRegion>
                <cx:axis id="0">
                <cx:catScaling gapWidth="0.06"/>
                <cx:tickLabels/>
                </cx:axis>
                </cx:plotArea>
                </cx:chart>
                </cx:chartSpace>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_chartex_histogram() {
        let mut values = ChartRange::new_from_range("Sheet1", 0, 0, 4, 0);
        values.set_cache(
            &["17", "25", "42", "63", "91"],
            ChartRangeCacheDataType::Number,
        );

        let mut chart = Chart::new(ChartType::Histogram);
        chart.add_series().set_values(&values);

        chart
            .set_histogram_bin_width(20.0)
            .set_histogram_underflow(20.0)
            .set_histogram_overflow(80.0);

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);

        assert!(got.contains(
            "<cx:layoutPr>\
             <cx:binning intervalClosed=\"r\" underflow=\"20\" overflow=\"80\">\
             <cx:binSize val=\"20\"/>\
             </cx:binning>\
             </cx:layoutPr>"
        ));
        assert!(got.contains("<cx:series layoutId=\"clusteredColumn\""));
    }

    #[test]
    fn test_chartex_pareto() {
        let mut values = ChartRange::new_from_range("Sheet1", 0, 0, 4, 0);
        values.set_cache(
            &["17", "25", "42", "63", "91"],
            ChartRangeCacheDataType::Number,
        );

        let mut chart = Chart::new(ChartType::Pareto);
        chart.add_series().set_values(&values);

        chart.set_histogram_bin_count(5);

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);

        // The histogram part of the chart is associated with the primary
        // value axis.
        assert!(got.contains(
            "<cx:binning intervalClosed=\"r\">\
             <cx:binCount val=\"5\"/>\
             </cx:binning>\
             </cx:layoutPr>\
             <cx:axisId val=\"1\"/>\
             </cx:series>"
        ));

        // The cumulative percentage line is a second series associated with a
        // secondary percentage axis.
        assert!(got.contains(
            "<cx:series layoutId=\"paretoLine\" ownerIdx=\"0\" uniqueId=\"{00000000-0000-0000-0000-000000000002}\" formatIdx=\"1\">\
             <cx:axisId val=\"2\"/>\
             </cx:series>"
        ));
        assert!(got.contains(
            "<cx:axis id=\"2\">\
             <cx:valScaling max=\"1\" min=\"0\"/>\
             <cx:units unit=\"percentage\"/>\
             <cx:tickLabels/>\
             </cx:axis>"
        ));
    }

    #[test]
    fn test_chartex_box_whisker() {
        let mut values = ChartRange::new_from_range("Sheet1", 0, 0, 4, 0);
        values.set_cache(
            &["17", "25", "42", "63", "91"],
            ChartRangeCacheDataType::Number,
        );

        let mut chart = Chart::new(ChartType::BoxWhisker);
        chart.add_series().set_values(&values);

        chart
            .set_box_whisker_quartile_method(ChartQuartileMethod::Inclusive)
            .set_box_whisker_mean_line(true)
            .set_box_whisker_mean_marker(false)
            .set_box_whisker_outliers(false)
            .set_box_whisker_inner_points(true);

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);

        assert!(got.contains("<cx:series layoutId=\"boxWhisker\""));
        assert!(got.contains(
            "<cx:layoutPr>\
             <cx:parentLabelLayout val=\"banner\"/>\
             <cx:visibility meanLine=\"1\" meanMarker=\"0\" nonoutliers=\"1\" outliers=\"0\"/>\
             <cx:statistics quartileMethod=\"inclusive\"/>\
             </cx:layoutPr>"
        ));
    }

    #[test]
    fn test_chartex_treemap() {
        // A two level category hierarchy: parent categories in the first
        // column and child categories in the second column.
        let mut categories = ChartRange::new_from_range("Sheet1", 0, 0, 2, 1);
        categories.cache = ChartRangeCacheData {
            major_dim: 3,
            minor_dim: 2,
            cache_type: ChartRangeCacheDataType::MultiLevelString,
            data: vec![
                "North".to_string(),
                "Apple".to_string(),
                String::new(),
                "Pear".to_string(),
                "South".to_string(),
                "Fig".to_string(),
            ],
        };

        let mut values = ChartRange::new_from_range("Sheet1", 0, 2, 2, 2);
        values.set_cache(&["100", "50", "30"], ChartRangeCacheDataType::Number);

        let mut chart = Chart::new(ChartType::Treemap);
        chart
            .add_series()
            .set_categories(&categories)
            .set_values(&values);

        chart.set_treemap_parent_labels(ChartParentLabelLayout::Banner);

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);

        // The category levels are written leaf level first.
        assert!(got.contains(
            "<cx:strDim type=\"cat\">\
             <cx:f>Sheet1!$A$1:$B$3</cx:f>\
             <cx:lvl ptCount=\"3\">\
             <cx:pt idx=\"0\">Apple</cx:pt>\
             <cx:pt idx=\"1\">Pear</cx:pt>\
             <cx:pt idx=\"2\">Fig</cx:pt>\
             </cx:lvl>\
             <cx:lvl ptCount=\"3\">\
             <cx:pt idx=\"0\">North</cx:pt>\
             <cx:pt idx=\"2\">South</cx:pt>\
             </cx:lvl>\
             </cx:strDim>"
        ));

        // Treemap and Sunburst charts use a "size" numeric dimension.
        assert!(got.contains("<cx:numDim type=\"size\">"));

        assert!(got.contains("<cx:series layoutId=\"treemap\""));
        assert!(got.contains(
            "<cx:layoutPr>\
             <cx:parentLabelLayout val=\"banner\"/>\
             </cx:layoutPr>"
        ));

        // Treemap charts don't have axes.
        assert!(!got.contains("<cx:axis"));
    }

    #[test]
    fn test_chartex_sunburst() {
        let mut categories = ChartRange::new_from_range("Sheet1", 0, 0, 2, 0);
        categories.set_cache(&["Apple", "Pear", "Fig"], ChartRangeCacheDataType::String);

        let mut values = ChartRange::new_from_range("Sheet1", 0, 1, 2, 1);
        values.set_cache(&["100", "50", "30"], ChartRangeCacheDataType::Number);

        let mut chart = Chart::new(ChartType::Sunburst);
        chart
            .add_series()
            .set_categories(&categories)
            .set_values(&values)
            .set_format(
                ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FF0000")),
            );

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);

        assert!(got.contains("<cx:numDim type=\"size\">"));

        // The series fill is written to a cx:spPr element.
        assert!(got.contains(
            "<cx:series layoutId=\"sunburst\" uniqueId=\"{00000000-0000-0000-0000-000000000001}\" formatIdx=\"0\">\
             <cx:spPr>\
             <a:solidFill>\
             <a:srgbClr val=\"FF0000\"/>\
             </a:solidFill>\
             </cx:spPr>\
             <cx:dataId val=\"0\"/>\
             </cx:series>"
        ));
    }

    #[test]
    fn test_chartex_region_map() {
        let mut categories = ChartRange::new_from_range("Sheet1", 0, 0, 2, 0);
        categories.set_cache(
            &["France", "Germany", "Spain"],
            ChartRangeCacheDataType::String,
        );

        let mut values = ChartRange::new_from_range("Sheet1", 0, 1, 2, 1);
        values.set_cache(&["10", "20", "30"], ChartRangeCacheDataType::Number);

        let mut chart = Chart::new(ChartType::RegionMap);
        chart
            .add_series()
            .set_categories(&categories)
            .set_values(&values);

        chart
            .set_region_map_projection(ChartMapProjection::Mercator)
            .set_region_map_labels(ChartRegionLabelLayout::ShowAll);

        chart.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&chart.writer);

        // Region maps use a "colorVal" numeric dimension.
        assert!(got.contains("<cx:numDim type=\"colorVal\">"));

        assert!(got.contains("<cx:series layoutId=\"regionMap\""));
        assert!(got.contains(
            "<cx:layoutPr>\
             <cx:regionLabelLayout val=\"showAll\"/>\
             <cx:geography projectionType=\"mercator\" cultureLanguage=\"en-US\" cultureRegion=\"US\" attribution=\"Powered by Bing\"/>\
             </cx:layoutPr>"
        ));

        // Region maps don't have axes.
        assert!(!got.contains("<cx:axis"));
    }

    #[test]
    fn test_chartex_chartsheet_validation() {
        // ChartEx charts can't be added to chartsheets.
        let mut chart = Chart::new(ChartType::Waterfall);
        chart.add_series().set_values("Sheet1!$A$1:$A$3");
        chart.is_chartsheet = true;

        let result = chart.validate();
        assert!(matches!(result, Err(XlsxError::ChartError(_))));
    }
}
