// chart - A module for creating the Excel Chart.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use crate::{utility, xmlwriter::XMLWriter, ColNum, RowNum};

#[derive(Clone)]
pub struct Chart {
    pub(crate) writer: XMLWriter,

    axis_ids: (u32, u32),
    series: Vec<ChartSeries>,
}

/// TODO
#[allow(dead_code)]

impl Chart {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    // Create a new Chart struct.
    pub fn new() -> Chart {
        let writer = XMLWriter::new();

        Chart {
            writer,
            axis_ids: (0, 0),
            series: vec![],
        }
    }

    // TODO.
    pub fn add_series(mut self, series: &ChartSeries) -> Chart {
        self.series.push(series.clone());

        self
    }

    /// Set default values for the chart axis ids.
    ///
    /// This is mainly used to ensure that the axis ids used in testing match
    /// the semi-randomized values in the target Excel file.
    ///
    #[doc(hidden)]
    pub fn set_axis_ids(&mut self, axis_id1: u32, axis_id2: u32) {
        self.axis_ids = (axis_id1, axis_id2);
    }

    // -----------------------------------------------------------------------
    // Crate level helper methods.
    // -----------------------------------------------------------------------

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the c:chartSpace element.
        self.write_chart_space();

        // Write the c:lang element.
        self.write_lang();

        // Write the c:chart element.
        self.write_chart();

        // Write the c:printSettings element.
        self.write_print_settings();

        // Close the c:chartSpace tag.
        self.writer.xml_end_tag("c:chartSpace");
    }

    // Write the <c:chartSpace> element.
    fn write_chart_space(&mut self) {
        let attributes = vec![
            (
                "xmlns:c",
                "http://schemas.openxmlformats.org/drawingml/2006/chart".to_string(),
            ),
            (
                "xmlns:a",
                "http://schemas.openxmlformats.org/drawingml/2006/main".to_string(),
            ),
            (
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string(),
            ),
        ];

        self.writer.xml_start_tag_attr("c:chartSpace", &attributes);
    }

    // Write the <c:lang> element.
    fn write_lang(&mut self) {
        let attributes = vec![("val", "en-US".to_string())];

        self.writer.xml_empty_tag_attr("c:lang", &attributes);
    }

    // Write the <c:chart> element.
    fn write_chart(&mut self) {
        self.writer.xml_start_tag("c:chart");

        // Write the c:plotArea element.
        self.write_plot_area();

        // Write the c:legend element.
        self.write_legend();

        // Write the c:plotVisOnly element.
        self.write_plot_vis_only();

        self.writer.xml_end_tag("c:chart");
    }

    // Write the <c:plotArea> element.
    fn write_plot_area(&mut self) {
        self.writer.xml_start_tag("c:plotArea");

        // Write the c:layout element.
        self.write_layout();

        // Write the c:barChart element.
        self.write_bar_chart();

        // Write the c:catAx element.
        self.write_cat_ax();

        // Write the c:valAx element.
        self.write_val_ax();

        self.writer.xml_end_tag("c:plotArea");
    }

    // Write the <c:layout> element.
    fn write_layout(&mut self) {
        self.writer.xml_empty_tag("c:layout");
    }

    // Write the <c:barChart> element.
    fn write_bar_chart(&mut self) {
        self.writer.xml_start_tag("c:barChart");

        // Write the c:barDir element.
        self.write_bar_dir();

        // Write the c:grouping element.
        self.write_grouping();

        // Write the c:ser elements.
        self.write_series();

        // Write the c:axId elements.
        self.write_ax_ids();

        self.writer.xml_end_tag("c:barChart");
    }

    // Write the <c:barDir> element.
    fn write_bar_dir(&mut self) {
        let attributes = vec![("val", "bar".to_string())];

        self.writer.xml_empty_tag_attr("c:barDir", &attributes);
    }

    // Write the <c:grouping> element.
    fn write_grouping(&mut self) {
        let attributes = vec![("val", "clustered".to_string())];

        self.writer.xml_empty_tag_attr("c:grouping", &attributes);
    }

    // Write the <c:ser> element.
    fn write_series(&mut self) {
        for (index, series) in self.series.clone().iter().enumerate() {
            self.writer.xml_start_tag("c:ser");

            // Write the c:idx element.
            self.write_idx(index);

            // Write the c:order element.
            self.write_order(index);

            // Write the c:cat element.
            self.write_cat(&series.category_range, &series.category_cache_data);

            // Write the c:val element.
            self.write_val(&series.value_range, &series.value_cache_data);

            self.writer.xml_end_tag("c:ser");
        }
    }

    // Write the <c:idx> element.
    fn write_idx(&mut self, index: usize) {
        let attributes = vec![("val", index.to_string())];

        self.writer.xml_empty_tag_attr("c:idx", &attributes);
    }

    // Write the <c:order> element.
    fn write_order(&mut self, index: usize) {
        let attributes = vec![("val", index.to_string())];

        self.writer.xml_empty_tag_attr("c:order", &attributes);
    }

    // Write the <c:cat> element.
    fn write_cat(&mut self, range: &ChartRange, cache: &[String]) {
        self.writer.xml_start_tag("c:cat");

        // Write the c:numRef element.
        self.write_num_ref(range, cache);

        self.writer.xml_end_tag("c:cat");
    }

    // Write the <c:val> element.
    fn write_val(&mut self, range: &ChartRange, cache: &[String]) {
        self.writer.xml_start_tag("c:val");

        // Write the c:numRef element.
        self.write_num_ref(range, cache);

        self.writer.xml_end_tag("c:val");
    }

    // Write the <c:numRef> element.
    fn write_num_ref(&mut self, range: &ChartRange, cache: &[String]) {
        self.writer.xml_start_tag("c:numRef");

        // Write the c:f element.
        self.write_range_formula(&range.formula());

        // Write the c:numCache element.
        self.write_num_cache(cache);

        self.writer.xml_end_tag("c:numRef");
    }

    // Write the <c:f> element.
    fn write_range_formula(&mut self, formula: &str) {
        self.writer.xml_data_element("c:f", formula);
    }

    // Write the <c:numCache> element.
    fn write_num_cache(&mut self, cache: &[String]) {
        self.writer.xml_start_tag("c:numCache");

        // Write the c:formatCode element.
        self.write_format_code();

        // Write the c:ptCount element.
        self.write_pt_count(cache.len());

        // Write the c:pt elements.
        for (index, value) in cache.iter().enumerate() {
            self.write_pt(index, value);
        }

        self.writer.xml_end_tag("c:numCache");
    }

    // Write the <c:formatCode> element.
    fn write_format_code(&mut self) {
        self.writer.xml_data_element("c:formatCode", "General");
    }

    // Write the <c:ptCount> element.
    fn write_pt_count(&mut self, count: usize) {
        let attributes = vec![("val", count.to_string())];

        self.writer.xml_empty_tag_attr("c:ptCount", &attributes);
    }

    // Write the <c:pt> element.
    fn write_pt(&mut self, index: usize, value: &str) {
        let attributes = vec![("idx", index.to_string())];

        self.writer.xml_start_tag_attr("c:pt", &attributes);
        self.writer.xml_data_element("c:v", value);
        self.writer.xml_end_tag("c:pt");
    }

    // Write both <c:axId> elements.
    fn write_ax_ids(&mut self) {
        self.write_ax_id(self.axis_ids.0);
        self.write_ax_id(self.axis_ids.1);
    }

    // Write the <c:axId> element.
    fn write_ax_id(&mut self, axis_id: u32) {
        let attributes = vec![("val", axis_id.to_string())];

        self.writer.xml_empty_tag_attr("c:axId", &attributes);
    }

    // Write the <c:catAx> element.
    fn write_cat_ax(&mut self) {
        self.writer.xml_start_tag("c:catAx");

        self.write_ax_id(self.axis_ids.0);

        // Write the c:scaling element.
        self.write_scaling();

        // Write the c:axPos element.
        self.write_ax_pos("l");

        // Write the c:numFmt element.
        self.write_num_fmt();

        // Write the c:tickLblPos element.
        self.write_tick_lbl_pos();

        // Write the c:crossAx element.
        self.write_cross_ax(self.axis_ids.1);

        // Write the c:crosses element.
        self.write_crosses();

        // Write the c:auto element.
        self.write_auto();

        // Write the c:lblAlgn element.
        self.write_lbl_algn();

        // Write the c:lblOffset element.
        self.write_lbl_offset();

        self.writer.xml_end_tag("c:catAx");
    }

    // Write the <c:valAx> element.
    fn write_val_ax(&mut self) {
        self.writer.xml_start_tag("c:valAx");

        self.write_ax_id(self.axis_ids.1);

        // Write the c:scaling element.
        self.write_scaling();

        // Write the c:axPos element.
        self.write_ax_pos("b");

        // Write the c:majorGridlines element.
        self.write_major_gridlines();

        // Write the c:numFmt element.
        self.write_num_fmt();

        // Write the c:tickLblPos element.
        self.write_tick_lbl_pos();

        // Write the c:crossAx element.
        self.write_cross_ax(self.axis_ids.0);

        // Write the c:crosses element.
        self.write_crosses();

        // Write the c:crossBetween element.
        self.write_cross_between();

        self.writer.xml_end_tag("c:valAx");
    }

    // Write the <c:scaling> element.
    fn write_scaling(&mut self) {
        self.writer.xml_start_tag("c:scaling");

        // Write the c:orientation element.
        self.write_orientation();

        self.writer.xml_end_tag("c:scaling");
    }

    // Write the <c:orientation> element.
    fn write_orientation(&mut self) {
        let attributes = vec![("val", "minMax".to_string())];

        self.writer.xml_empty_tag_attr("c:orientation", &attributes);
    }

    // Write the <c:axPos> element.
    fn write_ax_pos(&mut self, position: &str) {
        let attributes = vec![("val", position.to_string())];

        self.writer.xml_empty_tag_attr("c:axPos", &attributes);
    }

    // Write the <c:numFmt> element.
    fn write_num_fmt(&mut self) {
        let attributes = vec![
            ("formatCode", "General".to_string()),
            ("sourceLinked", "1".to_string()),
        ];

        self.writer.xml_empty_tag_attr("c:numFmt", &attributes);
    }

    // Write the <c:majorGridlines> element.
    fn write_major_gridlines(&mut self) {
        self.writer.xml_empty_tag("c:majorGridlines");
    }

    // Write the <c:tickLblPos> element.
    fn write_tick_lbl_pos(&mut self) {
        let attributes = vec![("val", "nextTo".to_string())];

        self.writer.xml_empty_tag_attr("c:tickLblPos", &attributes);
    }

    // Write the <c:crossAx> element.
    fn write_cross_ax(&mut self, axis_id: u32) {
        let attributes = vec![("val", axis_id.to_string())];

        self.writer.xml_empty_tag_attr("c:crossAx", &attributes);
    }

    // Write the <c:crosses> element.
    fn write_crosses(&mut self) {
        let attributes = vec![("val", "autoZero".to_string())];

        self.writer.xml_empty_tag_attr("c:crosses", &attributes);
    }

    // Write the <c:auto> element.
    fn write_auto(&mut self) {
        let attributes = vec![("val", "1".to_string())];

        self.writer.xml_empty_tag_attr("c:auto", &attributes);
    }

    // Write the <c:lblAlgn> element.
    fn write_lbl_algn(&mut self) {
        let attributes = vec![("val", "ctr".to_string())];

        self.writer.xml_empty_tag_attr("c:lblAlgn", &attributes);
    }

    // Write the <c:lblOffset> element.
    fn write_lbl_offset(&mut self) {
        let attributes = vec![("val", "100".to_string())];

        self.writer.xml_empty_tag_attr("c:lblOffset", &attributes);
    }

    // Write the <c:crossBetween> element.
    fn write_cross_between(&mut self) {
        let attributes = vec![("val", "between".to_string())];

        self.writer
            .xml_empty_tag_attr("c:crossBetween", &attributes);
    }

    // Write the <c:legend> element.
    fn write_legend(&mut self) {
        self.writer.xml_start_tag("c:legend");

        // Write the c:legendPos element.
        self.write_legend_pos();

        // Write the c:layout element.
        self.write_layout();

        self.writer.xml_end_tag("c:legend");
    }

    // Write the <c:legendPos> element.
    fn write_legend_pos(&mut self) {
        let attributes = vec![("val", "r".to_string())];

        self.writer.xml_empty_tag_attr("c:legendPos", &attributes);
    }

    // Write the <c:plotVisOnly> element.
    fn write_plot_vis_only(&mut self) {
        let attributes = vec![("val", "1".to_string())];

        self.writer.xml_empty_tag_attr("c:plotVisOnly", &attributes);
    }

    // Write the <c:printSettings> element.
    fn write_print_settings(&mut self) {
        self.writer.xml_start_tag("c:printSettings");

        // Write the c:headerFooter element.
        self.write_header_footer();

        // Write the c:pageMargins element.
        self.write_page_margins();

        // Write the c:pageSetup element.
        self.write_page_setup();

        self.writer.xml_end_tag("c:printSettings");
    }

    // Write the <c:headerFooter> element.
    fn write_header_footer(&mut self) {
        self.writer.xml_empty_tag("c:headerFooter");
    }

    // Write the <c:pageMargins> element.
    fn write_page_margins(&mut self) {
        let attributes = vec![
            ("b", "0.75".to_string()),
            ("l", "0.7".to_string()),
            ("r", "0.7".to_string()),
            ("t", "0.75".to_string()),
            ("header", "0.3".to_string()),
            ("footer", "0.3".to_string()),
        ];

        self.writer.xml_empty_tag_attr("c:pageMargins", &attributes);
    }

    // Write the <c:pageSetup> element.
    fn write_page_setup(&mut self) {
        self.writer.xml_empty_tag("c:pageSetup");
    }
}

// -----------------------------------------------------------------------
// Secondary structs.
// -----------------------------------------------------------------------

/// TODO
#[allow(dead_code)]
#[derive(Clone)]
pub struct ChartSeries {
    value_range: ChartRange,
    category_range: ChartRange,
    value_cache_is_numeric: bool,
    category_cache_is_numeric: bool,
    value_cache_data: Vec<String>,
    category_cache_data: Vec<String>,
}

impl ChartSeries {
    pub fn new() -> ChartSeries {
        ChartSeries {
            value_range: ChartRange::new("", 0, 0, 0, 0),
            category_range: ChartRange::new("", 0, 0, 0, 0),
            value_cache_is_numeric: true,
            category_cache_is_numeric: true,
            value_cache_data: vec![],
            category_cache_data: vec![],
        }
    }
    pub fn set_values(
        mut self,
        sheet_name: &str,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
    ) -> ChartSeries {
        self.value_range = ChartRange::new(sheet_name, first_row, first_col, last_row, last_col);
        self
    }

    pub fn set_categories(
        mut self,
        sheet_name: &str,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
    ) -> ChartSeries {
        self.category_range = ChartRange::new(sheet_name, first_row, first_col, last_row, last_col);
        self
    }

    pub fn set_value_cache(mut self, data: &[String], is_numeric: bool) -> ChartSeries {
        self.value_cache_data = data.to_vec();
        self.value_cache_is_numeric = is_numeric;
        self
    }

    pub fn set_category_cache(mut self, data: &[String], is_numeric: bool) -> ChartSeries {
        self.category_cache_data = data.to_vec();
        self.category_cache_is_numeric = is_numeric;
        self
    }
}

/// TODO
#[allow(dead_code)]
#[derive(Clone)]
struct ChartRange {
    sheet_name: String,
    first_row: RowNum,
    first_col: ColNum,
    last_row: RowNum,
    last_col: ColNum,
}

impl ChartRange {
    fn new(
        sheet_name: &str,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
    ) -> ChartRange {
        ChartRange {
            sheet_name: sheet_name.to_string(),
            first_row,
            first_col,
            last_row,
            last_col,
        }
    }

    fn formula(&self) -> String {
        utility::chart_range_abs(
            &self.sheet_name,
            self.first_row,
            self.first_col,
            self.last_row,
            self.last_col,
        )
    }
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use crate::chart::Chart;
    use crate::test_functions::xml_to_vec;
    use crate::ChartSeries;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let series1 = ChartSeries::new()
            .set_categories("Sheet1", 0, 0, 4, 0)
            .set_values("Sheet1", 0, 1, 4, 1)
            .set_category_cache(
                &[
                    "1".to_string(),
                    "2".to_string(),
                    "3".to_string(),
                    "4".to_string(),
                    "5".to_string(),
                ],
                true,
            )
            .set_value_cache(
                &[
                    "2".to_string(),
                    "4".to_string(),
                    "6".to_string(),
                    "8".to_string(),
                    "10".to_string(),
                ],
                true,
            );

        let series2 = ChartSeries::new()
            .set_categories("Sheet1", 0, 0, 4, 0)
            .set_values("Sheet1", 0, 2, 4, 2)
            .set_category_cache(
                &[
                    "1".to_string(),
                    "2".to_string(),
                    "3".to_string(),
                    "4".to_string(),
                    "5".to_string(),
                ],
                true,
            )
            .set_value_cache(
                &[
                    "3".to_string(),
                    "6".to_string(),
                    "9".to_string(),
                    "12".to_string(),
                    "15".to_string(),
                ],
                true,
            );

        let mut chart = Chart::new().add_series(&series1).add_series(&series2);

        chart.set_axis_ids(64052224, 64055552);

        chart.assemble_xml_file();

        let got = chart.writer.read_to_str();
        let got = xml_to_vec(&got);

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
}
