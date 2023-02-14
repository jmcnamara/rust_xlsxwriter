// chart - A module for creating the Excel Chart.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use crate::{
    drawing::{DrawingObject, DrawingType},
    utility,
    xmlwriter::XMLWriter,
    ColNum, ObjectMovement, RowNum,
};

// TODO remove all the dead_code attributes.

#[derive(Clone)]
#[allow(dead_code)] // TODO
pub struct Chart {
    pub(crate) id: u32,
    pub(crate) writer: XMLWriter,
    pub(crate) x_offset: u32,
    pub(crate) y_offset: u32,
    pub(crate) alt_text: String,
    pub(crate) vml_name: String,
    pub(crate) object_movement: ObjectMovement,
    pub(crate) decorative: bool,
    pub(crate) drawing_type: DrawingType,
    pub(crate) series: Vec<ChartSeries>,
    height: f64,
    width: f64,
    scale_width: f64,
    scale_height: f64,
    axis_ids: (u32, u32),
    category_has_num_format: bool,
    chart_type: ChartType,
    chart_group_type: ChartType,
    x_axis: ChartAxis,
    y_axis: ChartAxis,
    grouping: ChartGrouping,
    default_cross_between: bool,
    default_num_format: String,
    has_overlap: bool,
    overlap: i8,
}

/// TODO
#[allow(dead_code)] // TODO
impl Chart {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    // Create a new Chart struct.
    #[allow(clippy::new_without_default)]
    pub fn new(chart_type: ChartType) -> Chart {
        let writer = XMLWriter::new();

        let chart = Chart {
            writer,
            id: 0,
            height: 288.0,
            width: 480.0,
            scale_width: 1.0,
            scale_height: 1.0,
            x_offset: 0,
            y_offset: 0,
            alt_text: "".to_string(),
            vml_name: "image".to_string(),
            object_movement: ObjectMovement::MoveAndSizeWithCells,
            decorative: false,
            drawing_type: DrawingType::Chart,

            axis_ids: (0, 0),
            series: vec![],
            category_has_num_format: false,
            chart_type,
            chart_group_type: chart_type,
            x_axis: ChartAxis::new(),
            y_axis: ChartAxis::new(),
            grouping: ChartGrouping::Standard,
            default_cross_between: true,
            default_num_format: "General".to_string(),
            has_overlap: false,
            overlap: 0,
        };

        match chart_type {
            ChartType::Area | ChartType::AreaStacked | ChartType::AreaPercentStacked => {
                Self::initialize_area_chart(chart)
            }
            ChartType::Bar | ChartType::BarStacked | ChartType::BarPercentStacked => {
                Self::initialize_bar_chart(chart)
            }
            ChartType::Column | ChartType::ColumnStacked | ChartType::ColumnPercentStacked => {
                Self::initialize_column_chart(chart)
            }
            ChartType::Doughnut => Self::initialize_doughnut_chart(chart),
            ChartType::Line => Self::initialize_line_chart(chart),
            ChartType::Pie => Self::initialize_pie_chart(chart),
            ChartType::Radar => Self::initialize_radar_chart(chart),
            ChartType::Scatter
            | ChartType::ScatterStraight
            | ChartType::ScatterStraightWithMarkers
            | ChartType::ScatterSmooth
            | ChartType::ScatterSmoothWithMarkers => Self::initialize_scatter_chart(chart),
        }
    }

    // TODO.
    pub fn add_series(&mut self) -> &mut ChartSeries {
        let series = ChartSeries::new();
        self.series.push(series);

        self.series.last_mut().unwrap()
    }

    // TODO.
    pub fn push_series(&mut self, series: &ChartSeries) -> &mut Chart {
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

    // Set chart unique axis ids.
    pub(crate) fn add_axis_ids(&mut self) {
        if self.axis_ids.0 != 0 {
            return;
        }

        let axis_id_1 = (5000 + self.id) * 10000 + 1;
        let axis_id_2 = axis_id_1 + 1;

        self.axis_ids = (axis_id_1, axis_id_2);
    }

    // -----------------------------------------------------------------------
    // Chart specific methods.
    // -----------------------------------------------------------------------

    // Initialize area charts.
    fn initialize_area_chart(mut self) -> Chart {
        self.x_axis.axis_type = ChartAxisType::Category;
        self.x_axis.axis_position = ChartAxisPosition::Bottom;

        self.y_axis.axis_type = ChartAxisType::Value;
        self.y_axis.axis_position = ChartAxisPosition::Left;

        self.chart_group_type = ChartType::Area;
        self.default_cross_between = false;

        if self.chart_type == ChartType::Area {
            self.grouping = ChartGrouping::Standard;
        } else if self.chart_type == ChartType::AreaStacked {
            self.grouping = ChartGrouping::Stacked;
        } else if self.chart_type == ChartType::AreaPercentStacked {
            self.grouping = ChartGrouping::PercentStacked;
            self.default_num_format = "0%".to_string();
        }

        self
    }

    // Initialize bar charts. Bar chart category/value axes are reversed in
    // comparison to other charts. Some of the defaults reflect this.
    fn initialize_bar_chart(mut self) -> Chart {
        self.x_axis.axis_type = ChartAxisType::Value;
        self.x_axis.axis_position = ChartAxisPosition::Bottom;

        self.y_axis.axis_type = ChartAxisType::Category;
        self.y_axis.axis_position = ChartAxisPosition::Left;

        self.chart_group_type = ChartType::Bar;

        if self.chart_type == ChartType::Bar {
            self.grouping = ChartGrouping::Clustered;
        } else if self.chart_type == ChartType::BarStacked {
            self.grouping = ChartGrouping::Stacked;
            self.has_overlap = true;
            self.overlap = 100;
        } else if self.chart_type == ChartType::BarPercentStacked {
            self.grouping = ChartGrouping::PercentStacked;
            self.default_num_format = "0%".to_string();
            self.has_overlap = true;
            self.overlap = 100;
        }

        self
    }

    // Initialize column charts.
    fn initialize_column_chart(mut self) -> Chart {
        self.x_axis.axis_type = ChartAxisType::Category;
        self.x_axis.axis_position = ChartAxisPosition::Bottom;

        self.y_axis.axis_type = ChartAxisType::Value;
        self.y_axis.axis_position = ChartAxisPosition::Left;

        self.chart_group_type = ChartType::Column;

        if self.chart_type == ChartType::Column {
            self.grouping = ChartGrouping::Clustered;
        } else if self.chart_type == ChartType::ColumnStacked {
            self.grouping = ChartGrouping::Stacked;
            self.has_overlap = true;
            self.overlap = 100;
        } else if self.chart_type == ChartType::ColumnPercentStacked {
            self.grouping = ChartGrouping::PercentStacked;
            self.default_num_format = "0%".to_string();
            self.has_overlap = true;
            self.overlap = 100;
        }

        self
    }

    // Initialize doughnut charts.
    fn initialize_doughnut_chart(mut self) -> Chart {
        self.chart_group_type = ChartType::Doughnut;

        self
    }

    // Initialize line charts.
    fn initialize_line_chart(mut self) -> Chart {
        self.x_axis.axis_type = ChartAxisType::Category;
        self.x_axis.axis_position = ChartAxisPosition::Bottom;

        self.y_axis.axis_type = ChartAxisType::Value;
        self.y_axis.axis_position = ChartAxisPosition::Left;

        self.chart_group_type = ChartType::Line;

        self
    }

    // Initialize pie charts.
    fn initialize_pie_chart(mut self) -> Chart {
        self.chart_group_type = ChartType::Pie;

        self
    }

    // Initialize radar charts.
    fn initialize_radar_chart(mut self) -> Chart {
        self.x_axis.axis_type = ChartAxisType::Category;
        self.x_axis.axis_position = ChartAxisPosition::Bottom;

        self.y_axis.axis_type = ChartAxisType::Value;
        self.y_axis.axis_position = ChartAxisPosition::Left;

        self.chart_group_type = ChartType::Radar;

        self
    }

    // Initialize scatter charts.
    fn initialize_scatter_chart(mut self) -> Chart {
        self.x_axis.axis_type = ChartAxisType::Category;
        self.x_axis.axis_position = ChartAxisPosition::Bottom;

        self.y_axis.axis_type = ChartAxisType::Value;
        self.y_axis.axis_position = ChartAxisPosition::Left;

        self.chart_group_type = ChartType::Scatter;
        self.default_cross_between = false;

        self
    }

    // Write the <c:areaChart> element for Column charts.
    fn write_area_chart(&mut self) {
        self.writer.xml_start_tag("c:areaChart");

        // Write the c:grouping element.
        self.write_grouping();

        // Write the c:ser elements.
        self.write_series();

        // Write the c:axId elements.
        self.write_ax_ids();

        self.writer.xml_end_tag("c:areaChart");
    }

    // Write the <c:barChart> element for Bar charts.
    fn write_bar_chart(&mut self) {
        self.writer.xml_start_tag("c:barChart");

        // Write the c:barDir element.
        self.write_bar_dir("bar");

        // Write the c:grouping element.
        self.write_grouping();

        // Write the c:ser elements.
        self.write_series();

        if self.has_overlap {
            // Write the c:overlap element.
            self.write_overlap();
        }

        // Write the c:axId elements.
        self.write_ax_ids();

        self.writer.xml_end_tag("c:barChart");
    }

    // Write the <c:barChart> element for Column charts.
    fn write_column_chart(&mut self) {
        self.writer.xml_start_tag("c:barChart");

        // Write the c:barDir element.
        self.write_bar_dir("col");

        // Write the c:grouping element.
        self.write_grouping();

        // Write the c:ser elements.
        self.write_series();

        if self.has_overlap {
            // Write the c:overlap element.
            self.write_overlap();
        }

        // Write the c:axId elements.
        self.write_ax_ids();

        self.writer.xml_end_tag("c:barChart");
    }

    // Write the <c:doughnutChart> element for Column charts.
    fn write_doughnut_chart(&mut self) {
        self.writer.xml_start_tag("c:doughnutChart");

        // Write the c:varyColors element.
        self.write_vary_colors();

        // Write the c:ser elements.
        self.write_series();

        // Write the c:firstSliceAng element.
        self.write_first_slice_ang();

        // Write the c:holeSize element.
        self.write_hole_size();

        self.writer.xml_end_tag("c:doughnutChart");
    }

    // Write the <c:lineChart>element.
    fn write_line_chart(&mut self) {
        self.writer.xml_start_tag("c:lineChart");

        // Write the c:grouping element.
        self.write_grouping();

        // Write the c:ser elements.
        self.write_series();

        // Write the c:marker element.
        self.write_marker_value();

        // Write the c:axId elements.
        self.write_ax_ids();

        self.writer.xml_end_tag("c:lineChart");
    }

    // Write the <c:pieChart> element for Column charts.
    fn write_pie_chart(&mut self) {
        self.writer.xml_start_tag("c:pieChart");

        // Write the c:varyColors element.
        self.write_vary_colors();

        // Write the c:ser elements.
        self.write_series();

        // Write the c:firstSliceAng element.
        self.write_first_slice_ang();

        self.writer.xml_end_tag("c:pieChart");
    }

    // Write the <c:radarChart>element.
    fn write_radar_chart(&mut self) {
        self.writer.xml_start_tag("c:radarChart");

        // Write the c:radarStyle element.
        self.write_radar_style();

        // Write the c:ser elements.
        self.write_series();

        // Write the c:axId elements.
        self.write_ax_ids();

        self.writer.xml_end_tag("c:radarChart");
    }

    // Write the <c:scatterChart>element.
    fn write_scatter_chart(&mut self) {
        self.writer.xml_start_tag("c:scatterChart");

        // Write the c:scatterStyle element.
        self.write_scatter_style();

        // Write the c:ser elements.
        self.write_scatter_series();

        // Write the c:axId elements.
        self.write_ax_ids();

        self.writer.xml_end_tag("c:scatterChart");
    }

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

        match self.chart_type {
            ChartType::Area | ChartType::AreaStacked | ChartType::AreaPercentStacked => {
                self.write_area_chart()
            }
            ChartType::Bar | ChartType::BarStacked | ChartType::BarPercentStacked => {
                self.write_bar_chart()
            }
            ChartType::Column | ChartType::ColumnStacked | ChartType::ColumnPercentStacked => {
                self.write_column_chart()
            }
            ChartType::Doughnut => self.write_doughnut_chart(),
            ChartType::Line => self.write_line_chart(),
            ChartType::Pie => self.write_pie_chart(),
            ChartType::Radar => self.write_radar_chart(),
            ChartType::Scatter
            | ChartType::ScatterStraight
            | ChartType::ScatterStraightWithMarkers
            | ChartType::ScatterSmooth
            | ChartType::ScatterSmoothWithMarkers => self.write_scatter_chart(),
        }

        // Reverse the X and Y axes for Bar charts.
        if self.chart_group_type == ChartType::Bar {
            std::mem::swap(&mut self.x_axis, &mut self.y_axis);
        }

        match self.chart_group_type {
            ChartType::Pie | ChartType::Doughnut => {}

            ChartType::Scatter => {
                // Write the c:valAx element.
                self.write_cat_val_ax();

                // Write the c:valAx element.
                self.write_val_ax();
            }
            _ => {
                // Write the c:catAx element.
                self.write_cat_ax();

                // Write the c:valAx element.
                self.write_val_ax();
            }
        }

        // Reset the X and Y axes for Bar charts.
        if self.chart_group_type == ChartType::Bar {
            std::mem::swap(&mut self.x_axis, &mut self.y_axis);
        }

        self.writer.xml_end_tag("c:plotArea");
    }

    // Write the <c:layout> element.
    fn write_layout(&mut self) {
        self.writer.xml_empty_tag("c:layout");
    }

    // Write the <c:barDir> element.
    fn write_bar_dir(&mut self, direction: &str) {
        let attributes = vec![("val", direction.to_string())];

        self.writer.xml_empty_tag_attr("c:barDir", &attributes);
    }

    // Write the <c:grouping> element.
    fn write_grouping(&mut self) {
        let attributes = vec![("val", self.grouping.to_string())];

        self.writer.xml_empty_tag_attr("c:grouping", &attributes);
    }

    // Write the <c:scatterStyle> element.
    fn write_scatter_style(&mut self) {
        let mut attributes = vec![];

        if self.chart_type == ChartType::ScatterSmooth
            || self.chart_type == ChartType::ScatterSmoothWithMarkers
        {
            attributes.push(("val", "smoothMarker".to_string()))
        } else {
            attributes.push(("val", "lineMarker".to_string()))
        }

        self.writer
            .xml_empty_tag_attr("c:scatterStyle", &attributes);
    }

    // Write the <c:ser> element.
    fn write_series(&mut self) {
        for (index, series) in self.series.clone().iter().enumerate() {
            self.writer.xml_start_tag("c:ser");

            // Write the c:idx element.
            self.write_idx(index);

            // Write the c:order element.
            self.write_order(index);

            // Write the c:marker element.
            if self.chart_type == ChartType::Line || self.chart_type == ChartType::Radar {
                self.write_marker();
            }

            // Write the c:cat element.
            if series.category_range.has_data() {
                self.category_has_num_format = true;
                self.write_cat(&series.category_range, &series.category_cache_data);
            }

            // Write the c:val element.
            self.write_val(&series.value_range, &series.value_cache_data);

            self.writer.xml_end_tag("c:ser");
        }
    }

    // Write the <c:ser> element for scatter charts.
    fn write_scatter_series(&mut self) {
        for (index, series) in self.series.clone().iter().enumerate() {
            self.writer.xml_start_tag("c:ser");

            // Write the c:idx element.
            self.write_idx(index);

            // Write the c:order element.
            self.write_order(index);

            if self.chart_type == ChartType::ScatterStraight
                || self.chart_type == ChartType::ScatterSmooth
            {
                self.write_marker();
            }

            if self.chart_type == ChartType::Scatter {
                // Write the c:spPr element.
                self.write_sp_pr();
            }

            self.write_x_val(&series.category_range, &series.category_cache_data);

            self.write_y_val(&series.value_range, &series.value_cache_data);

            if self.chart_type == ChartType::ScatterSmooth
                || self.chart_type == ChartType::ScatterSmoothWithMarkers
            {
                // Write the c:smooth element.
                self.write_smooth();
            }

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
    fn write_cat(&mut self, range: &ChartRange, cache: &ChartSeriesCacheData) {
        self.writer.xml_start_tag("c:cat");

        // Write the c:numRef element.
        self.write_num_ref(range, cache);

        self.writer.xml_end_tag("c:cat");
    }

    // Write the <c:val> element.
    fn write_val(&mut self, range: &ChartRange, cache: &ChartSeriesCacheData) {
        self.writer.xml_start_tag("c:val");

        // Write the c:numRef element.
        self.write_num_ref(range, cache);

        self.writer.xml_end_tag("c:val");
    }

    // Write the <c:xVal> element for scatter charts.
    fn write_x_val(&mut self, range: &ChartRange, cache: &ChartSeriesCacheData) {
        self.writer.xml_start_tag("c:xVal");

        // Write the c:numRef element.
        self.write_num_ref(range, cache);

        self.writer.xml_end_tag("c:xVal");
    }

    // Write the <c:yVal> element for scatter charts.
    fn write_y_val(&mut self, range: &ChartRange, cache: &ChartSeriesCacheData) {
        self.writer.xml_start_tag("c:yVal");

        // Write the c:numRef element.
        self.write_num_ref(range, cache);

        self.writer.xml_end_tag("c:yVal");
    }

    // Write the <c:numRef> element.
    fn write_num_ref(&mut self, range: &ChartRange, cache: &ChartSeriesCacheData) {
        self.writer.xml_start_tag("c:numRef");

        // Write the c:f element.
        self.write_range_formula(&range.formula());

        // Write the c:numCache element.
        if cache.has_data() {
            self.write_num_cache(cache);
        }

        self.writer.xml_end_tag("c:numRef");
    }

    // Write the <c:f> element.
    fn write_range_formula(&mut self, formula: &str) {
        self.writer.xml_data_element("c:f", formula);
    }

    // Write the <c:numCache> element.
    fn write_num_cache(&mut self, cache: &ChartSeriesCacheData) {
        self.writer.xml_start_tag("c:numCache");

        // Write the c:formatCode element.
        self.write_format_code();

        // Write the c:ptCount element.
        self.write_pt_count(cache.data.len());

        // Write the c:pt elements.
        for (index, value) in cache.data.iter().enumerate() {
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
        self.write_ax_pos(self.x_axis.axis_position);

        if self.chart_type == ChartType::Radar {
            self.write_major_gridlines();
        }

        // Write the c:numFmt element.
        if self.category_has_num_format {
            self.write_category_num_fmt();
        }

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
        self.write_ax_pos(self.y_axis.axis_position);

        // Write the c:majorGridlines element.
        self.write_major_gridlines();

        // Write the c:numFmt element.
        self.write_value_num_fmt();

        // Write the c:majorTickMark element.
        if self.chart_type == ChartType::Radar {
            self.write_major_tick_mark();
        }

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

    // Write the category <c:valAx> element for scatter charts.
    fn write_cat_val_ax(&mut self) {
        self.writer.xml_start_tag("c:valAx");

        self.write_ax_id(self.axis_ids.0);

        // Write the c:scaling element.
        self.write_scaling();

        // Write the c:axPos element.
        self.write_ax_pos(self.x_axis.axis_position);

        // Write the c:numFmt element.
        self.write_value_num_fmt();

        // Write the c:tickLblPos element.
        self.write_tick_lbl_pos();

        // Write the c:crossAx element.
        self.write_cross_ax(self.axis_ids.1);

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
    fn write_ax_pos(&mut self, position: ChartAxisPosition) {
        let attributes = vec![("val", position.to_string())];

        self.writer.xml_empty_tag_attr("c:axPos", &attributes);
    }

    // Write the <c:numFmt> element.
    fn write_category_num_fmt(&mut self) {
        let attributes = vec![
            ("formatCode", "General".to_string()),
            ("sourceLinked", "1".to_string()),
        ];

        self.writer.xml_empty_tag_attr("c:numFmt", &attributes);
    }

    // Write the <c:numFmt> element.
    fn write_value_num_fmt(&mut self) {
        let attributes = vec![
            ("formatCode", self.default_num_format.clone()),
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
        let mut attributes = vec![];

        if self.default_cross_between {
            attributes.push(("val", "between".to_string()));
        } else {
            attributes.push(("val", "midCat".to_string()));
        }

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

        if self.chart_type == ChartType::Pie || self.chart_type == ChartType::Doughnut {
            // Write the c:txPr element.
            self.write_pie_tx_pr();
        }

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

    // Write the <c:marker> element.
    fn write_marker_value(&mut self) {
        let attributes = vec![("val", "1".to_string())];

        self.writer.xml_empty_tag_attr("c:marker", &attributes);
    }

    // Write the <c:marker> element.
    fn write_marker(&mut self) {
        self.writer.xml_start_tag("c:marker");

        // Write the c:symbol element.
        self.write_symbol();

        self.writer.xml_end_tag("c:marker");
    }

    // Write the <c:symbol> element.
    fn write_symbol(&mut self) {
        let attributes = vec![("val", "none".to_string())];

        self.writer.xml_empty_tag_attr("c:symbol", &attributes);
    }

    // Write the <c:varyColors> element.
    fn write_vary_colors(&mut self) {
        let attributes = vec![("val", "1".to_string())];

        self.writer.xml_empty_tag_attr("c:varyColors", &attributes);
    }

    // Write the <c:firstSliceAng> element.
    fn write_first_slice_ang(&mut self) {
        let attributes = vec![("val", "0".to_string())];

        self.writer
            .xml_empty_tag_attr("c:firstSliceAng", &attributes);
    }

    // Write the <c:holeSize> element.
    fn write_hole_size(&mut self) {
        let attributes = vec![("val", "50".to_string())];

        self.writer.xml_empty_tag_attr("c:holeSize", &attributes);
    }

    // Write the <c:txPr> element.
    fn write_pie_tx_pr(&mut self) {
        self.writer.xml_start_tag("c:txPr");

        // Write the a:bodyPr element.
        self.write_a_body_pr();

        // Write the a:lstStyle element.
        self.write_a_lst_style();

        // Write the a:p element.
        self.write_a_p();

        self.writer.xml_end_tag("c:txPr");
    }

    // Write the <a:bodyPr> element.
    fn write_a_body_pr(&mut self) {
        self.writer.xml_empty_tag("a:bodyPr");
    }

    // Write the <a:lstStyle> element.
    fn write_a_lst_style(&mut self) {
        self.writer.xml_empty_tag("a:lstStyle");
    }

    // Write the <a:p> element.
    fn write_a_p(&mut self) {
        self.writer.xml_start_tag("a:p");

        // Write the a:pPr element.
        self.write_pie_a_p_pr();

        // Write the a:endParaRPr element.
        self.write_a_end_para_rpr();

        self.writer.xml_end_tag("a:p");
    }

    // Write the <a:pPr> element.
    fn write_pie_a_p_pr(&mut self) {
        let attributes = vec![("rtl", "0".to_string())];

        self.writer.xml_start_tag_attr("a:pPr", &attributes);

        // Write the a:defRPr element.
        self.write_a_def_rpr();
        self.writer.xml_end_tag("a:pPr");
    }

    // Write the <a:defRPr> element.
    fn write_a_def_rpr(&mut self) {
        self.writer.xml_empty_tag("a:defRPr");
    }

    // Write the <a:endParaRPr> element.
    fn write_a_end_para_rpr(&mut self) {
        let attributes = vec![("lang", "en-US".to_string())];

        self.writer.xml_empty_tag_attr("a:endParaRPr", &attributes);
    }

    // Write the <c:spPr> element.
    fn write_sp_pr(&mut self) {
        self.writer.xml_start_tag("c:spPr");

        // Write the a:ln element.
        self.write_a_ln();

        self.writer.xml_end_tag("c:spPr");
    }

    // Write the <a:ln> element.
    fn write_a_ln(&mut self) {
        let attributes = vec![("w", "28575".to_string())];

        self.writer.xml_start_tag_attr("a:ln", &attributes);

        // Write the a:noFill element.
        self.write_a_no_fill();
        self.writer.xml_end_tag("a:ln");
    }

    // Write the <a:noFill> element.
    fn write_a_no_fill(&mut self) {
        self.writer.xml_empty_tag("a:noFill");
    }

    // Write the <c:radarStyle> element.
    fn write_radar_style(&mut self) {
        let attributes = vec![("val", "marker".to_string())];

        self.writer.xml_empty_tag_attr("c:radarStyle", &attributes);
    }

    // Write the <c:majorTickMark> element.
    fn write_major_tick_mark(&mut self) {
        let attributes = vec![("val", "cross".to_string())];

        self.writer
            .xml_empty_tag_attr("c:majorTickMark", &attributes);
    }

    // Write the <c:overlap> element.
    fn write_overlap(&mut self) {
        let attributes = vec![("val", "100".to_string())];

        self.writer.xml_empty_tag_attr("c:overlap", &attributes);
    }

    // Write the <c:smooth> element.
    fn write_smooth(&mut self) {
        let attributes = vec![("val", "1".to_string())];

        self.writer.xml_empty_tag_attr("c:smooth", &attributes);
    }
}

// -----------------------------------------------------------------------
// Traits.
// -----------------------------------------------------------------------

// TODO
impl DrawingObject for Chart {
    fn x_offset(&self) -> u32 {
        self.x_offset
    }

    fn y_offset(&self) -> u32 {
        self.y_offset
    }

    fn width_scaled(&self) -> f64 {
        self.width * self.scale_width
    }

    fn height_scaled(&self) -> f64 {
        self.height * self.scale_height
    }

    fn object_movement(&self) -> ObjectMovement {
        self.object_movement
    }

    fn alt_text(&self) -> String {
        self.alt_text.clone()
    }

    fn decorative(&self) -> bool {
        self.decorative
    }

    fn drawing_type(&self) -> DrawingType {
        self.drawing_type
    }
}

// -----------------------------------------------------------------------
// Secondary structs.
// -----------------------------------------------------------------------

/// TODO
#[allow(dead_code)] // todo
#[derive(Clone)]
pub struct ChartSeries {
    pub(crate) value_range: ChartRange,
    pub(crate) category_range: ChartRange,
    pub(crate) value_cache_data: ChartSeriesCacheData,
    pub(crate) category_cache_data: ChartSeriesCacheData,
}

#[allow(clippy::new_without_default)]
impl ChartSeries {
    pub fn new() -> ChartSeries {
        ChartSeries {
            value_range: ChartRange::new("", 0, 0, 0, 0),
            category_range: ChartRange::new("", 0, 0, 0, 0),
            value_cache_data: ChartSeriesCacheData::new(),
            category_cache_data: ChartSeriesCacheData::new(),
        }
    }
    pub fn set_values(
        &mut self,
        sheet_name: &str,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
    ) -> &mut ChartSeries {
        self.value_range = ChartRange::new(sheet_name, first_row, first_col, last_row, last_col);
        self
    }

    pub fn set_categories(
        &mut self,
        sheet_name: &str,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
    ) -> &mut ChartSeries {
        self.category_range = ChartRange::new(sheet_name, first_row, first_col, last_row, last_col);
        self
    }

    pub fn set_value_cache(&mut self, data: &[&str], is_numeric: bool) -> &mut ChartSeries {
        self.value_cache_data = ChartSeriesCacheData {
            is_numeric,
            data: data.iter().map(|s| s.to_string()).collect(),
        };
        self
    }

    pub fn set_category_cache(&mut self, data: &[&str], is_numeric: bool) -> &mut ChartSeries {
        self.category_cache_data = ChartSeriesCacheData {
            is_numeric,
            data: data.iter().map(|s| s.to_string()).collect(),
        };
        self
    }
}

/// TODO
#[allow(dead_code)]
#[derive(Clone)]
pub(crate) struct ChartRange {
    sheet_name: String,
    first_row: RowNum,
    first_col: ColNum,
    last_row: RowNum,
    last_col: ColNum,
}

impl ChartRange {
    pub(crate) fn new(
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

    pub(crate) fn formula(&self) -> String {
        utility::chart_range_abs(
            &self.sheet_name,
            self.first_row,
            self.first_col,
            self.last_row,
            self.last_col,
        )
    }

    pub(crate) fn key(&self) -> (String, RowNum, ColNum, RowNum, ColNum) {
        (
            self.sheet_name.clone(),
            self.first_row,
            self.first_col,
            self.last_row,
            self.last_col,
        )
    }

    pub(crate) fn has_data(&self) -> bool {
        !self.sheet_name.is_empty()
    }
}

#[allow(dead_code)]
#[derive(Clone)]
pub(crate) struct ChartSeriesCacheData {
    pub(crate) is_numeric: bool,
    pub(crate) data: Vec<String>,
}

impl ChartSeriesCacheData {
    pub(crate) fn new() -> ChartSeriesCacheData {
        ChartSeriesCacheData {
            is_numeric: false,
            data: vec![],
        }
    }

    pub(crate) fn has_data(&self) -> bool {
        !self.data.is_empty()
    }
}

#[allow(dead_code)]
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ChartType {
    Area,
    AreaStacked,
    AreaPercentStacked,

    Bar,
    BarStacked,
    BarPercentStacked,

    Column,
    ColumnStacked,
    ColumnPercentStacked,

    Doughnut,

    Line,

    Pie,

    Radar,

    Scatter,
    ScatterStraight,
    ScatterStraightWithMarkers,
    ScatterSmooth,
    ScatterSmoothWithMarkers,
}

#[allow(dead_code)]
#[derive(Clone)]
pub(crate) struct ChartAxis {
    axis_type: ChartAxisType,
    axis_position: ChartAxisPosition,
}

impl ChartAxis {
    pub(crate) fn new() -> ChartAxis {
        ChartAxis {
            axis_type: ChartAxisType::Value,
            axis_position: ChartAxisPosition::Bottom,
        }
    }
}

#[allow(dead_code)]
#[derive(Clone)]
pub(crate) enum ChartAxisType {
    Category,
    Value,
}

#[allow(dead_code)]
#[derive(Clone, Copy)]
pub(crate) enum ChartAxisPosition {
    Bottom,
    Left,
}

impl ToString for ChartAxisPosition {
    fn to_string(&self) -> String {
        match self {
            ChartAxisPosition::Bottom => "b".to_string(),
            ChartAxisPosition::Left => "l".to_string(),
        }
    }
}

#[allow(dead_code)]
#[derive(Clone, Copy)]
pub(crate) enum ChartGrouping {
    Stacked,
    Standard,
    Clustered,
    PercentStacked,
}

impl ToString for ChartGrouping {
    fn to_string(&self) -> String {
        match self {
            ChartGrouping::Stacked => "stacked".to_string(),
            ChartGrouping::Standard => "standard".to_string(),
            ChartGrouping::Clustered => "clustered".to_string(),
            ChartGrouping::PercentStacked => "percentStacked".to_string(),
        }
    }
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use crate::chart::{Chart, ChartType};
    use crate::test_functions::xml_to_vec;
    use crate::ChartSeries;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let mut series1 = ChartSeries::new();
        series1
            .set_categories("Sheet1", 0, 0, 4, 0)
            .set_values("Sheet1", 0, 1, 4, 1)
            .set_category_cache(&["1", "2", "3", "4", "5"], true)
            .set_value_cache(&["2", "4", "6", "8", "10"], true);

        let mut series2 = ChartSeries::new();
        series2
            .set_categories("Sheet1", 0, 0, 4, 0)
            .set_values("Sheet1", 0, 2, 4, 2)
            .set_category_cache(&["1", "2", "3", "4", "5"], true)
            .set_value_cache(&["3", "6", "9", "12", "15"], true);

        let mut chart = Chart::new(ChartType::Bar);
        chart.push_series(&series1).push_series(&series2);

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
