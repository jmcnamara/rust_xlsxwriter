// chart unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod chart_tests {

    use crate::chart::{Chart, ChartRange, ChartSeries, ChartType, XlsxError};
    use crate::test_functions::xml_to_vec;
    use crate::ChartRangeCacheDataType;
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

        let got = chart.writer.read_to_str();
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

    #[test]
    fn test_range_from_string() {
        let range_string = "=Sheet1!$A$1:$A$5";
        let range = ChartRange::new_from_string(range_string);
        assert_eq!("Sheet1!$A$1:$A$5", range.formula_abs());
        assert_eq!("Sheet1", range.sheet_name);

        let range_string = "Sheet1!$A$1:$A$5";
        let range = ChartRange::new_from_string(range_string);
        assert_eq!("Sheet1!$A$1:$A$5", range.formula_abs());
        assert_eq!("Sheet1", range.sheet_name);

        let range_string = "Sheet 1!$A$1:$A$5";
        let range = ChartRange::new_from_string(range_string);
        assert_eq!("'Sheet 1'!$A$1:$A$5", range.formula_abs());
        assert_eq!("Sheet 1", range.sheet_name);
    }
}
