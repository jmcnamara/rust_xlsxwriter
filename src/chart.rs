// chart - A module for creating the Excel Chart.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use regex::Regex;

use crate::{
    drawing::{DrawingObject, DrawingType},
    utility,
    xmlwriter::XMLWriter,
    ColNum, ObjectMovement, RowNum,
};

#[derive(Clone)]
/// The Chart struct is used to create an object to represent an chart that can
/// be inserted into a worksheet.
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_intro.png">
///
/// Code the generate the above file:
///
/// ```
/// # // This code is available in examples/doc_chart_intro.rs
/// #
/// use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///     let worksheet = workbook.add_worksheet();
///
///     // Add some test data for the charts.
///     worksheet.write(0, 0, 10)?;
///     worksheet.write(1, 0, 60)?;
///     worksheet.write(2, 0, 30)?;
///     worksheet.write(3, 0, 10)?;
///     worksheet.write(4, 0, 50)?;
///
///     // Create a simple Column chart.
///     let mut chart = Chart::new(ChartType::Column);
///
///     // Add a data series using Excel formula syntax to describe the range.
///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
///
///    // Add the chart to the worksheet.
///     worksheet.insert_chart(0, 2, &chart)?;
///
///     // Save the file.
///     workbook.save("chart.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// # Chart Value and Category Axes
///
/// When working with charts it is important to understand how Excel
/// differentiates between a chart axis that is used for series categories and a
/// chart axis that is used for series values.
///
/// In the majority of Excel charts the X axis is the **category** axis and each
/// of the values is evenly spaced and sequential. The Y axis is the **value**
/// axis and points are displayed according to their value:
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_axes01.png">
///
/// Excel treats these two types of axis differently and exposes different
/// properties for each. For example, here are the properties for a category
/// axis:
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_axes02.png">
///
/// Here are properties for a value axis:
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_axes03.png">
///
/// As such, some of the `rust_xlsxwriter` axis properties can be set for a
/// value axis, some can be set for a category axis and some properties can be
/// set for both. For example `reverse` can be set for either category or value
/// axes while the `min` and `max` properties can only be set for value axes
/// (and date axes). The documentation calls out the type of axis to which
/// properties apply.
///
/// For a Bar chart the Category and Value axes are reversed:
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_axes04.png">
///
/// A Scatter chart (but not a Line chart) has 2 value axes:
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_axes05.png">
///
/// Date category axes are a special type of category axis that give them some
/// of the properties of values axes such as `min` and `max` when used with date
/// or time values. These aren't currently supported but will be in a future
/// release.
///
pub struct Chart {
    pub(crate) id: u32,
    pub(crate) writer: XMLWriter,
    pub(crate) x_offset: u32,
    pub(crate) y_offset: u32,
    pub(crate) alt_text: String,
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
    pub(crate) title: ChartTitle,
    pub(crate) x_axis: ChartAxis,
    pub(crate) y_axis: ChartAxis,
    grouping: ChartGrouping,
    default_cross_between: bool,
    default_num_format: String,
    has_overlap: bool,
    overlap: i8,
    style: u8,
}

impl Chart {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new Chart struct.
    ///
    /// Create a new [`Chart`] object that can be configured and inserted into a
    /// worksheet using the
    /// [worksheet.insert_chart()][crate::Worksheet::insert_chart].
    ///
    /// # Examples
    ///
    /// A simple chart example using the rust_xlsxwriter library.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_simple.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 50)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #
    /// #     // Create a simple Column chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_simple.png">
    ///
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
            object_movement: ObjectMovement::MoveAndSizeWithCells,
            decorative: false,
            drawing_type: DrawingType::Chart,

            axis_ids: (0, 0),
            series: vec![],
            category_has_num_format: false,
            chart_type,
            chart_group_type: chart_type,
            title: ChartTitle::new(),
            x_axis: ChartAxis::new(),
            y_axis: ChartAxis::new(),
            grouping: ChartGrouping::Standard,
            default_cross_between: true,
            default_num_format: "General".to_string(),
            has_overlap: false,
            overlap: 0,
            style: 2,
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

            ChartType::Line | ChartType::LineStacked | ChartType::LinePercentStacked => {
                Self::initialize_line_chart(chart)
            }

            ChartType::Pie => Self::initialize_pie_chart(chart),

            ChartType::Radar | ChartType::RadarWithMarkers | ChartType::RadarFilled => {
                Self::initialize_radar_chart(chart)
            }

            ChartType::Scatter
            | ChartType::ScatterStraight
            | ChartType::ScatterStraightWithMarkers
            | ChartType::ScatterSmooth
            | ChartType::ScatterSmoothWithMarkers => Self::initialize_scatter_chart(chart),
        }
    }

    /// Add a new chart series to a chart.
    ///
    /// Create and add a new chart series to a chart. The chart series
    /// represents the Category and Value ranges as well as formatting and
    /// display options. A chart in Excel must contain one of more data series.
    /// A series is represented by a [`ChartSeries`] struct.
    ///
    /// A chart series is usually created via the `chart.add_series()`] method.
    /// However, if required you can create a standalone `ChartSeries` object
    /// and add it to a chart via the
    /// [`chart.push_series()`](Chart::push_series) method, see below.
    ///
    /// # Examples
    ///
    /// An example of creating a chart series via `chart.add_series()`.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_add_series.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 50)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #
    /// #     // Create a simple Column chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_simple.png">
    ///
    pub fn add_series(&mut self) -> &mut ChartSeries {
        let series = ChartSeries::new();
        self.series.push(series);

        self.series.last_mut().unwrap()
    }

    /// Add a new chart series to a chart.
    ///
    /// Create and add a new chart series to a chart. The chart series
    /// represents the Category and Value ranges as well as formatting and
    /// display options. A chart in Excel must contain one of more data series.
    /// A series is represented by a [`ChartSeries`] struct.
    ///
    /// A chart series is usually created via the
    /// [`chart.add_series()`](Chart::add_series) method, see above. However, if
    /// required you can create a standalone `ChartSeries` object and add it to
    /// a chart via the `chart.push_series()` method.
    ///
    /// # Examples
    ///
    /// An example of creating a chart series as a standalone object and then
    /// adding it to a chart via the `chart.push_series()` method.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_push_series.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, ChartSeries, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 50)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #
    /// #     // Create a simple Column chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Create a chart series and set the range for the values.
    ///     let mut series = ChartSeries::new();
    ///     series.set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Add the data series to the chart.
    ///     chart.push_series(&series);
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file for both examples:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_simple.png">
    ///
    pub fn push_series(&mut self, series: &ChartSeries) -> &mut Chart {
        self.series.push(series.clone());
        self
    }

    /// Get the chart title object in order to set its properties.
    ///
    /// Get a reference to the charts X-Axis [`ChartTitle`] object in order to
    /// set its properties.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting properties of the chart title.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_title_set_name.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 50)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #
    /// #     // Create a simple Column chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Set the chart title.
    ///     chart.title().set_name("This is the chart title");
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_title_set_name.png">
    ///
    pub fn title(&mut self) -> &mut ChartTitle {
        &mut self.title
    }

    /// Get the chart X-Axis object in order to set its properties.
    ///
    /// Get a reference to the charts X-Axis [`ChartAxis`] object in order to
    /// set its properties.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting properties of the axes.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_name.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 50)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #
    /// #     // Create a simple Column chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Set the chart axis titles.
    ///     chart.x_axis().set_name("Test number");
    ///     chart.y_axis().set_name("Sample length (mm)");
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_name.png">
    pub fn x_axis(&mut self) -> &mut ChartAxis {
        &mut self.x_axis
    }

    /// Get the chart Y-Axis object in order to set its properties.
    ///
    /// Get a reference to the charts Y-Axis [`ChartAxis`] object in order to
    /// set its properties.
    ///
    /// See the [`chart.x_axis()`][Chart::x_axis] method above.
    ///
    pub fn y_axis(&mut self) -> &mut ChartAxis {
        &mut self.y_axis
    }

    /// Set the chart style type.
    ///
    /// The `set_style()` method is used to set the style of the chart to one of
    /// 48 built-in styles.
    ///
    /// These styles were available in the original Excel 2007 interface. In
    /// later versions they have been replaced with "layouts" on the "Chart
    /// Design" tab. These layouts are not defined in the file format. They are
    /// a collection of modifications to the base chart type. They can be
    /// replicated using the Chart APIs (when complete) but they cannot be defined by
    /// the `set_style()` method.
    ///
    /// # Arguments
    ///
    /// * `style` - A integer value in the range 1-48.
    ///
    /// # Examples
    ///
    /// # An example showing all 48 default chart styles available in Excel 2007
    /// using rust_xlsxwriter. Note, these styles are not the same as the styles
    /// available in Excel 2013.
    ///
    /// ```
    /// # // This code is available in examples/app_chart_styles.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     let chart_types = vec![
    ///         ("Column", ChartType::Column),
    ///         ("Area", ChartType::Area),
    ///         ("Line", ChartType::Line),
    ///         ("Pie", ChartType::Pie),
    ///     ];
    ///
    ///     // Create a worksheet with 48 charts in each of the available styles, for
    ///     // each of the chart types above.
    ///     for (name, chart_type) in chart_types {
    ///         let worksheet = workbook.add_worksheet().set_name(name)?.set_zoom(30);
    ///         let mut chart = Chart::new(chart_type);
    ///         chart.add_series().set_values("Data!$A$1:$A$6");
    ///         let mut style = 1;
    ///
    ///         for row_num in (0..90).step_by(15) {
    ///             for col_num in (0..64).step_by(8) {
    ///                 chart.set_style(style);
    ///                 chart.title().set_name(&format!("Style {style}"));
    ///                 worksheet.insert_chart(row_num as u32, col_num as u16, &chart)?;
    ///                 style += 1;
    ///             }
    ///         }
    ///     }
    ///
    /// #     // Create a worksheet with data for the charts.
    /// #     let data_worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     data_worksheet.write(0, 0, 10)?;
    /// #     data_worksheet.write(1, 0, 40)?;
    /// #     data_worksheet.write(2, 0, 50)?;
    /// #     data_worksheet.write(3, 0, 20)?;
    /// #     data_worksheet.write(4, 0, 10)?;
    /// #     data_worksheet.write(5, 0, 50)?;
    /// #     data_worksheet.set_hidden(true);
    /// #
    /// #     workbook.save("chart_styles.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_styles.png">
    ///
    pub fn set_style(&mut self, style: u8) -> &mut Chart {
        if (1..=48).contains(&style) {
            self.style = style;
        } else {
            eprintln!("Style id {style} outside Excel range: 1 <= style <= 48.");
        }

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

        self.y_axis.title.is_horizontal = true;

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

        self.y_axis.title.is_horizontal = true;

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

        self.y_axis.title.is_horizontal = true;

        self.chart_group_type = ChartType::Line;

        if self.chart_type == ChartType::Line {
            self.grouping = ChartGrouping::Standard;
        } else if self.chart_type == ChartType::LineStacked {
            self.grouping = ChartGrouping::Stacked;
        } else if self.chart_type == ChartType::LinePercentStacked {
            self.grouping = ChartGrouping::PercentStacked;
            self.default_num_format = "0%".to_string();
        }

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

        self.y_axis.title.is_horizontal = true;

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

        // Write the c:style element.
        if self.style != 2 {
            self.write_style();
        }

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

        // Write the c:title element.
        if self.title.hidden {
            self.write_auto_title_deleted();
        } else {
            self.write_chart_title(&self.title.clone());
        }

        // Write the c:plotArea element.
        self.write_plot_area();

        // Write the c:legend element.
        self.write_legend();

        // Write the c:plotVisOnly element.
        self.write_plot_vis_only();

        self.writer.xml_end_tag("c:chart");
    }

    // Write the <c:title> element.
    fn write_chart_title(&mut self, title: &ChartTitle) {
        if !title.name.is_empty() {
            self.write_title_rich(title);
        } else if title.range.has_data() {
            self.write_title_formula(title);
        }
    }

    // Write the <c:title> element.
    fn write_series_title(&mut self, title: &ChartTitle) {
        if !title.name.is_empty() {
            self.write_tx_value(title);
        } else if title.range.has_data() {
            self.write_tx_formula(title);
        }
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

            ChartType::Line | ChartType::LineStacked | ChartType::LinePercentStacked => {
                self.write_line_chart()
            }

            ChartType::Pie => self.write_pie_chart(),

            ChartType::Radar | ChartType::RadarWithMarkers | ChartType::RadarFilled => {
                self.write_radar_chart()
            }

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

            self.write_series_title(&series.title);

            // Write the c:marker element.
            if self.chart_group_type == ChartType::Line || self.chart_type == ChartType::Radar {
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

            self.write_series_title(&series.title);

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

        self.write_cache_ref(range, cache);

        self.writer.xml_end_tag("c:cat");
    }

    // Write the <c:val> element.
    fn write_val(&mut self, range: &ChartRange, cache: &ChartSeriesCacheData) {
        self.writer.xml_start_tag("c:val");

        self.write_cache_ref(range, cache);

        self.writer.xml_end_tag("c:val");
    }

    // Write the <c:xVal> element for scatter charts.
    fn write_x_val(&mut self, range: &ChartRange, cache: &ChartSeriesCacheData) {
        self.writer.xml_start_tag("c:xVal");

        self.write_cache_ref(range, cache);

        self.writer.xml_end_tag("c:xVal");
    }

    // Write the <c:yVal> element for scatter charts.
    fn write_y_val(&mut self, range: &ChartRange, cache: &ChartSeriesCacheData) {
        self.writer.xml_start_tag("c:yVal");

        self.write_cache_ref(range, cache);

        self.writer.xml_end_tag("c:yVal");
    }

    // Write the <c:numRef> or <c:strRef> elements.
    fn write_cache_ref(&mut self, range: &ChartRange, cache: &ChartSeriesCacheData) {
        if cache.is_numeric {
            self.write_num_ref(range, cache);
        } else {
            self.write_str_ref(range, cache);
        }
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

    // Write the <c:strRef> element.
    fn write_str_ref(&mut self, range: &ChartRange, cache: &ChartSeriesCacheData) {
        self.writer.xml_start_tag("c:strRef");

        // Write the c:f element.
        self.write_range_formula(&range.formula());

        // Write the c:strCache element.
        if cache.has_data() {
            self.write_str_cache(cache);
        }

        self.writer.xml_end_tag("c:strRef");
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

    // Write the <c:strCache> element.
    fn write_str_cache(&mut self, cache: &ChartSeriesCacheData) {
        self.writer.xml_start_tag("c:strCache");

        // Write the c:ptCount element.
        self.write_pt_count(cache.data.len());

        // Write the c:pt elements.
        for (index, value) in cache.data.iter().enumerate() {
            self.write_pt(index, value);
        }

        self.writer.xml_end_tag("c:strCache");
    }

    // Write the <c:f> element.
    fn write_range_formula(&mut self, formula: &str) {
        self.writer.xml_data_element("c:f", formula);
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

        if self.chart_group_type == ChartType::Radar {
            self.write_major_gridlines();
        }

        // Write the c:title element.
        self.write_chart_title(&self.x_axis.title.clone());

        // Write the c:numFmt element.
        if self.category_has_num_format {
            self.write_category_num_fmt();
        }

        // Write the c:tickLblPos element.
        self.write_tick_label_position();

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

        // Write the c:title element.
        self.write_chart_title(&self.y_axis.title.clone());

        // Write the c:numFmt element.
        self.write_value_num_fmt();

        // Write the c:majorTickMark element.
        if self.chart_group_type == ChartType::Radar {
            self.write_major_tick_mark();
        }

        // Write the c:tickLblPos element.
        self.write_tick_label_position();

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

        // Write the c:title element.
        self.write_chart_title(&self.x_axis.title.clone());

        // Write the c:numFmt element.
        self.write_value_num_fmt();

        // Write the c:tickLblPos element.
        self.write_tick_label_position();

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
    fn write_tick_label_position(&mut self) {
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
            self.write_tx_pr_pie();
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
    fn write_tx_pr_pie(&mut self) {
        self.writer.xml_start_tag("c:txPr");

        // Write the a:bodyPr element.
        self.write_a_body_pr(false);

        // Write the a:lstStyle element.
        self.write_a_lst_style();

        // Write the a:p element.
        self.write_a_p_pie();

        self.writer.xml_end_tag("c:txPr");
    }

    // Write the <c:txPr> element.
    fn write_tx_pr(&mut self, is_horizontal: bool) {
        self.writer.xml_start_tag("c:txPr");

        // Write the a:bodyPr element.
        self.write_a_body_pr(is_horizontal);

        // Write the a:lstStyle element.
        self.write_a_lst_style();

        // Write the a:p element.
        self.write_a_p_formula();

        self.writer.xml_end_tag("c:txPr");
    }

    // Write the <a:p> element.
    fn write_a_p_formula(&mut self) {
        self.writer.xml_start_tag("a:p");

        // Write the a:pPr element.
        self.write_a_p_pr();

        // Write the a:endParaRPr element.
        self.write_a_end_para_rpr();

        self.writer.xml_end_tag("a:p");
    }

    // Write the <a:pPr> element.
    fn write_a_p_pr(&mut self) {
        self.writer.xml_start_tag("a:pPr");

        // Write the a:defRPr element.
        self.write_a_def_rpr();

        self.writer.xml_end_tag("a:pPr");
    }

    // Write the <a:bodyPr> element.
    fn write_a_body_pr(&mut self, is_horizontal: bool) {
        let mut attributes = vec![];
        let mut rotation = 0;

        if is_horizontal {
            rotation = -5400000;
        }

        if rotation != 0 {
            attributes.push(("rot", rotation.to_string()));
            attributes.push(("vert", "horz".to_string()));
        }

        self.writer.xml_empty_tag_attr("a:bodyPr", &attributes);
    }

    // Write the <a:lstStyle> element.
    fn write_a_lst_style(&mut self) {
        self.writer.xml_empty_tag("a:lstStyle");
    }

    // Write the <a:p> element.
    fn write_a_p_pie(&mut self) {
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
        let mut attributes = vec![];

        if self.chart_type == ChartType::RadarFilled {
            attributes.push(("val", "filled".to_string()));
        } else {
            attributes.push(("val", "marker".to_string()));
        }

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

    // Write the <c:style> element.
    fn write_style(&mut self) {
        let attributes = vec![("val", self.style.to_string())];

        self.writer.xml_empty_tag_attr("c:style", &attributes);
    }

    // Write the <c:autoTitleDeleted> element.
    fn write_auto_title_deleted(&mut self) {
        let attributes = vec![("val", "1".to_string())];

        self.writer
            .xml_empty_tag_attr("c:autoTitleDeleted", &attributes);
    }

    // Write the <c:title> element.
    fn write_title_formula(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag("c:title");

        // Write the c:tx element.
        self.write_tx_formula(title);

        // Write the c:layout element.
        self.write_layout();

        // Write the c:txPr element.
        self.write_tx_pr(title.is_horizontal);

        self.writer.xml_end_tag("c:title");
    }

    // Write the <c:tx> element.
    fn write_tx_formula(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag("c:tx");

        // Title is always a string type.
        self.write_str_ref(&title.range, &title.cache_data);

        self.writer.xml_end_tag("c:tx");
    }

    // Write the <c:title> element.
    fn write_title_rich(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag("c:title");

        // Write the c:tx element.
        self.write_tx_rich(title);

        // Write the c:layout element.
        self.write_layout();

        self.writer.xml_end_tag("c:title");
    }

    // Write the <c:tx> element.
    fn write_tx_rich(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag("c:tx");

        // Write the c:rich element.
        self.write_rich(title);

        self.writer.xml_end_tag("c:tx");
    }

    // Write the <c:tx> element.
    fn write_tx_value(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag("c:tx");

        self.writer.xml_data_element("c:v", &title.name);

        self.writer.xml_end_tag("c:tx");
    }

    // Write the <c:rich> element.
    fn write_rich(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag("c:rich");

        // Write the a:bodyPr element.
        self.write_a_body_pr(title.is_horizontal);

        // Write the a:lstStyle element.
        self.write_a_lst_style();

        // Write the a:p element.
        self.write_a_p_rich(title);

        self.writer.xml_end_tag("c:rich");
    }

    // Write the <a:p> element.
    fn write_a_p_rich(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag("a:p");

        // Write the a:pPr element.
        self.write_a_p_pr_rich();

        // Write the a:r element.
        self.write_a_r(title);

        self.writer.xml_end_tag("a:p");
    }

    // Write the <a:pPr> element.
    fn write_a_p_pr_rich(&mut self) {
        self.writer.xml_start_tag("a:pPr");

        // Write the a:defRPr element.
        self.write_a_def_rpr();

        self.writer.xml_end_tag("a:pPr");
    }

    // Write the <a:r> element.
    fn write_a_r(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag("a:r");

        // Write the a:rPr element.
        self.write_a_r_pr();

        // Write the a:t element.
        self.write_a_t(&title.name);

        self.writer.xml_end_tag("a:r");
    }

    // Write the <a:rPr> element.
    fn write_a_r_pr(&mut self) {
        let attributes = vec![("lang", "en-US".to_string())];

        self.writer.xml_empty_tag_attr("a:rPr", &attributes);
    }

    // Write the <a:t> element.
    fn write_a_t(&mut self, name: &str) {
        self.writer.xml_data_element("a:t", name);
    }
}

// -----------------------------------------------------------------------
// Traits.
// -----------------------------------------------------------------------

/// Trait to map types into an Excel chart range.
///
/// The 2 most common types of range used in rust_xlsxwriter charts are:
///
/// - A string with an Excel like range formula such as `"Sheet1!$A$1:$A$3"`.
/// - A 5 value tuple that can be used to create the range programmatically
///   using a sheet name and zero indexed row and column values like:
///   `("Sheet1", 0, 0, 2, 0)` (this gives the same range as the previous string
///   value).
///
/// For single cell ranges used in chart items such as chart or axis titles you
/// can also use:
///
/// - A simple string title.
/// - A string with an Excel like cell formula such as `"Sheet1!$A$1"`.
/// - A 3 value tuple that can be used to create the cell range programmatically
///   using a sheet name and zero indexed row and column values like:
///   `("Sheet1", 0, 0)` (this gives the same range as the previous string
///   value).
///
pub trait IntoChartRange {
    /// Trait function to turn a type into [`ChartRange`].
    fn new_chart_range(&self) -> ChartRange;
}

impl IntoChartRange for (&str, RowNum, ColNum, RowNum, ColNum) {
    fn new_chart_range(&self) -> ChartRange {
        ChartRange::new_from_range(self.0, self.1, self.2, self.3, self.4)
    }
}

impl IntoChartRange for (&str, RowNum, ColNum) {
    fn new_chart_range(&self) -> ChartRange {
        ChartRange::new_from_range(self.0, self.1, self.2, self.1, self.2)
    }
}

impl IntoChartRange for &str {
    fn new_chart_range(&self) -> ChartRange {
        ChartRange::new_from_string(self)
    }
}

impl IntoChartRange for &String {
    fn new_chart_range(&self) -> ChartRange {
        ChartRange::new_from_string(self)
    }
}

// Trait for objects that have a component stored in the drawing.xml file.
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

/// A struct to represent a Chart series.
///
/// A chart in Excel can contain one of more data series. The `ChartSeries`
/// struct represents the Category and Value ranges, and the formatting and
/// options for the chart series.
///
///
/// # Examples
///
/// A simple chart example using the rust_xlsxwriter library.
///
/// ```
/// // This code is available in examples/doc_chart_simple.rs
///
/// use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///     let worksheet = workbook.add_worksheet();
///
///     // Add some data for the chart.
///     worksheet.write(0, 0, 50)?;
///     worksheet.write(1, 0, 30)?;
///     worksheet.write(2, 0, 40)?;
///
///     // Create a simple Column chart.
///     let mut chart = Chart::new(ChartType::Column);
///
///     // Add a data series using Excel formula syntax to describe the range.
///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
///
///     // Add the chart to the worksheet.
///     worksheet.insert_chart(0, 2, &chart)?;
///
///     // Save the file.
///     workbook.save("chart.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_simple.png">
///
#[derive(Clone)]
pub struct ChartSeries {
    pub(crate) value_range: ChartRange,
    pub(crate) category_range: ChartRange,
    pub(crate) value_cache_data: ChartSeriesCacheData,
    pub(crate) category_cache_data: ChartSeriesCacheData,
    pub(crate) title: ChartTitle,
}

#[allow(clippy::new_without_default)]
impl ChartSeries {
    /// Create a new chart series object.
    ///
    /// Create a new chart series object. A chart in Excel can contain one of
    /// more data series. The `ChartSeries` struct represents the Category and
    /// Value ranges, and the formatting and options for the chart series.
    ///
    /// A chart series is usually created via the
    /// [`chart.add_series()`](Chart::add_series) method, see the first example
    /// below. However, if required you can create a standalone `ChartSeries`
    /// object and add it to a chart via the
    /// [`chart.push_series()`](Chart::push_series) method, see the second
    /// example below.
    ///
    /// # Examples
    ///
    /// An example of creating a chart series via
    /// [`chart.add_series()`](Chart::add_series).
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_add_series.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 50)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #
    /// #     // Create a simple Column chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// An example of creating a chart series as a standalone object and then adding
    /// it to a chart via the [`chart.push_series()`](Chart::add_series) method.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_push_series.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, ChartSeries, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 50)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #
    /// #     // Create a simple Column chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Create a chart series and set the range for the values.
    ///     let mut series = ChartSeries::new();
    ///     series.set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Add the data series to the chart.
    ///     chart.push_series(&series);
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file for both examples:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_simple.png">
    ///
    pub fn new() -> ChartSeries {
        ChartSeries {
            value_range: ChartRange::new_from_range("", 0, 0, 0, 0),
            category_range: ChartRange::new_from_range("", 0, 0, 0, 0),
            value_cache_data: ChartSeriesCacheData::new(),
            category_cache_data: ChartSeriesCacheData::new(),
            title: ChartTitle::new(),
        }
    }

    /// Add a values range to a chart series.
    ///
    /// All chart series in Excel must have a data range that defines the range
    /// of values for the series. In Excel this is typically a range like
    /// `"Sheet1!$B$1:$B$5"`.
    ///
    /// This is the most important property of a series and is the only
    /// mandatory option for every chart object. This series values links the
    /// chart with the worksheet data that it displays. The data range can be
    /// set using a formula as shown in the first part of the example below or
    /// using a list of values as shown in the second part.
    ///
    /// # Arguments
    ///
    /// * `range` - The range property which can be one of two generic types:
    ///    - A string with an Excel like range formula such as
    ///      `"Sheet1!$A$1:$A$3"`.
    ///    - A tuple that can be used to create the range programmatically using
    ///      a sheet name and zero indexed row and column values like:
    ///      `("Sheet1", 0, 0, 2, 0)` (this gives the same range as the previous
    ///      string value).
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the chart series values.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_series_set_values.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 50)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(0, 1, 30)?;
    /// #     worksheet.write(1, 1, 40)?;
    /// #     worksheet.write(2, 1, 50)?;
    /// #
    /// #     // Create a simple Column chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Add another data series to the chart using the alternative tuple syntax
    ///     // to describe the range. This method is better when you need to create the
    ///     // ranges programmatically to match the data range in the worksheet.
    ///     chart.add_series().set_values(("Sheet1", 0, 1, 2, 1));
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 3, &chart)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_series_set_values.png">
    ///
    pub fn set_values<T>(&mut self, range: T) -> &mut ChartSeries
    where
        T: IntoChartRange,
    {
        self.value_range = range.new_chart_range();
        self
    }

    /// Add a category range chart series.
    ///
    /// This method sets the chart category labels. The category is more or less
    /// the same as the X axis. In most chart types the categories property is
    /// optional and the chart will just assume a sequential series from `1..n`.
    /// The exception to this is the Scatter chart types for which a category
    /// range is mandatory in Excel.
    ///
    /// The data range can be set using a formula as shown in the first part of
    /// the example below or using a list of values as shown in the second part.
    ///
    /// # Arguments
    ///
    /// * `range` - The range property which can be one of two generic types:
    ///    - A string with an Excel like range formula such as
    ///      `"Sheet1!$A$1:$A$3"`.
    ///    - A tuple that can be used to create the range programmatically using
    ///      a sheet name and zero indexed row and column values like:
    ///      `("Sheet1", 0, 0, 2, 0)` (this gives the same range as the previous
    ///      string value).
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the chart series categories and
    /// values.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_series_set_categories.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, "Jan")?;
    /// #     worksheet.write(1, 0, "Feb")?;
    /// #     worksheet.write(2, 0, "Mar")?;
    /// #     worksheet.write(0, 1, 50)?;
    /// #     worksheet.write(1, 1, 30)?;
    /// #     worksheet.write(2, 1, 40)?;
    /// #     worksheet.write(0, 2, 30)?;
    /// #     worksheet.write(1, 2, 40)?;
    /// #     worksheet.write(2, 2, 50)?;
    /// #
    /// #     // Create a simple Column chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart
    ///         .add_series()
    ///         .set_categories("Sheet1!$A$1:$A$3")
    ///         .set_values("Sheet1!$B$1:$B$3");
    ///
    ///     // Add another data series to the chart using the alternative tuple syntax
    ///     // to describe the range. This method is better when you need to create the
    ///     // ranges programmatically to match the data range in the worksheet.
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 1, 2, 1))
    ///         .set_values(("Sheet1", 0, 2, 2, 2));
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 4, &chart)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_series_set_categories.png">
    ///
    pub fn set_categories<T>(&mut self, range: T) -> &mut ChartSeries
    where
        T: IntoChartRange,
    {
        self.category_range = range.new_chart_range();
        self
    }

    /// Add a name for a chart series.
    ///
    /// Set the name for the series. The name is displayed in the formula bar.
    /// For non-Pie/Doughnut charts it is also displayed in the legend. The name
    /// property is optional and if it isn’t supplied it will default to `Series
    /// 1..n`. The name can be a simple string, a formula such as `Sheet1!$A$1`
    /// or a tuple with a sheet name, row and column such as `('Sheet1', 0, 0)`.
    ///
    /// # Arguments
    ///
    /// * `range` - The range property which can be one of the following generic
    ///   types:
    ///    - A simple string title.
    ///    - A string with an Excel like range formula such as `"Sheet1!$A$1"`.
    ///    - A tuple that can be used to create the range programmatically using
    ///      a sheet name and zero indexed row and column values like:
    ///      `("Sheet1", 0, 0)` (this gives the same range as the previous
    ///      string value).
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the chart series name.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_series_set_name.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, "Month")?;
    /// #     worksheet.write(1, 0, "Jan")?;
    /// #     worksheet.write(2, 0, "Feb")?;
    /// #     worksheet.write(3, 0, "Mar")?;
    /// #     worksheet.write(0, 1, "Total")?;
    /// #     worksheet.write(1, 1, 30)?;
    /// #     worksheet.write(2, 1, 20)?;
    /// #     worksheet.write(3, 1, 40)?;
    /// #     worksheet.write(0, 2, "Q1")?;
    /// #     worksheet.write(1, 2, 10)?;
    /// #     worksheet.write(2, 2, 10)?;
    /// #     worksheet.write(3, 2, 10)?;
    /// #     worksheet.write(0, 3, "Q2")?;
    /// #     worksheet.write(1, 3, 20)?;
    /// #     worksheet.write(2, 3, 10)?;
    /// #     worksheet.write(3, 3, 30)?;
    /// #
    /// #     // Create a simple Column chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series with a simple string name.
    ///     chart
    ///         .add_series()
    ///         .set_name("Year to date")
    ///         .set_categories("Sheet1!$A$2:$A$4")
    ///         .set_values("Sheet1!$B$2:$B$4");
    ///
    ///
    ///     // Add a data series using Excel formula syntax to describe the range/name.
    ///     chart
    ///         .add_series()
    ///         .set_name("Sheet1!$C$1")
    ///         .set_categories("Sheet1!$A$2:$A$4")
    ///         .set_values("Sheet1!$C$2:$C$4");
    ///
    ///     // Add another data series to the chart using the alternative tuple syntax
    ///     // to describe the range/name. This method is better when you need to create
    ///     // the ranges programmatically to match the data range in the worksheet.
    ///     chart
    ///         .add_series()
    ///         .set_name(("Sheet1", 0, 3))
    ///         .set_categories(("Sheet1", 1, 0, 3, 0))
    ///         .set_values(("Sheet1", 1, 3, 3, 3));
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 5, &chart)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_series_set_name.png">
    ///
    pub fn set_name<T>(&mut self, name: T) -> &mut ChartSeries
    where
        T: IntoChartRange,
    {
        self.title.set_name(name);
        self
    }

    /// Add data to the chart values cache.
    ///
    /// This method is only used to populate the chart data caches in test code.
    /// The library reads and populates the cache automatically in most cases.
    #[doc(hidden)]
    pub fn set_value_cache(&mut self, data: &[&str], is_numeric: bool) -> &mut ChartSeries {
        self.value_cache_data = ChartSeriesCacheData {
            is_numeric,
            data: data.iter().map(|s| s.to_string()).collect(),
        };
        self
    }

    /// Add data to the chart categories cache.
    ///
    /// This method is only used to populate the chart data caches in test code.
    /// The library reads and populates the cache automatically in most cases.
    #[doc(hidden)]
    pub fn set_category_cache(&mut self, data: &[&str], is_numeric: bool) -> &mut ChartSeries {
        self.category_cache_data = ChartSeriesCacheData {
            is_numeric,
            data: data.iter().map(|s| s.to_string()).collect(),
        };
        self
    }
}

#[derive(Clone)]
/// A struct to represent a Chart range.
///
/// A struct to represent a chart range like `"Sheet1!$A$1:$A$4"`. The struct is
/// public to allow for the [`IntoChartRange`] trait but it isn't required to be
/// manipulated by the end user.
pub struct ChartRange {
    sheet_name: String,
    first_row: RowNum,
    first_col: ColNum,
    last_row: RowNum,
    last_col: ColNum,
    range_string: String,
}

impl ChartRange {
    pub(crate) fn new_from_range(
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
            range_string: "".to_string(),
        }
    }

    pub(crate) fn new_from_string(range_string: &str) -> ChartRange {
        lazy_static! {
            static ref CHART_CELL: Regex = Regex::new(r"^=?([^!]+)'?!\$?(\w+)\$?(\d+)").unwrap();
            static ref CHART_RANGE: Regex =
                Regex::new(r"^=?([^!]+)'?!\$?(\w+)\$?(\d+):\$?(\w+)\$?(\d+)").unwrap();
        }

        let mut sheet_name = "";
        let mut first_row = 0;
        let mut first_col = 0;
        let mut last_row = 0;
        let mut last_col = 0;

        if let Some(caps) = CHART_RANGE.captures(range_string) {
            sheet_name = caps.get(1).unwrap().as_str();
            first_row = caps.get(3).unwrap().as_str().parse::<u32>().unwrap() - 1;
            last_row = caps.get(5).unwrap().as_str().parse::<u32>().unwrap() - 1;
            first_col = utility::name_to_col(caps.get(2).unwrap().as_str());
            last_col = utility::name_to_col(caps.get(4).unwrap().as_str());
        } else if let Some(caps) = CHART_CELL.captures(range_string) {
            sheet_name = caps.get(1).unwrap().as_str();
            first_row = caps.get(3).unwrap().as_str().parse::<u32>().unwrap() - 1;
            first_col = utility::name_to_col(caps.get(2).unwrap().as_str());
            last_row = first_row;
            last_col = first_col;
        }

        let sheet_name: String = if sheet_name.starts_with('\'') && sheet_name.ends_with('\'') {
            sheet_name[1..sheet_name.len() - 1].to_string()
        } else {
            sheet_name.to_string()
        };

        ChartRange {
            sheet_name,
            first_row,
            first_col,
            last_row,
            last_col,
            range_string: range_string.to_string(),
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

#[derive(Clone)]
pub(crate) struct ChartSeriesCacheData {
    pub(crate) is_numeric: bool,
    pub(crate) data: Vec<String>,
}

impl ChartSeriesCacheData {
    pub(crate) fn new() -> ChartSeriesCacheData {
        ChartSeriesCacheData {
            is_numeric: true,
            data: vec![],
        }
    }

    pub(crate) fn has_data(&self) -> bool {
        !self.data.is_empty()
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
/// The ChartType enum define the type of a Chart object.
///
/// The main original chart types are supported, see below.
///
/// Stock chart variants will be supported at a later date. Support for newer
/// Excel chart types such as Treemap, Sunburst, Box and Whisker, Statistical
/// Histogram, Waterfall, Funnel and Maps is not currently planned.
///
pub enum ChartType {
    /// An Area chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_area1.png">
    Area,

    /// A stacked Area chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_area2.png">
    AreaStacked,

    /// A percent stacked Area chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_area3.png">
    AreaPercentStacked,

    /// A Bar (horizontal histogram) chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_bar1.png">
    Bar,

    /// A stacked Bar chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_bar2.png">
    BarStacked,

    /// A percent stacked Bar chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_bar3.png">
    BarPercentStacked,

    /// A Column (vertical histogram) chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_column1.png">
    Column,

    /// A stacked Column chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_column2.png">
    ColumnStacked,

    /// A percent stacked Column chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_column3.png">
    ColumnPercentStacked,

    /// A Doughnut chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_doughnut1.png">
    Doughnut,

    /// An Line chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line1.png">
    Line,

    /// A stacked Line chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line2.png">
    LineStacked,

    /// A percent stacked Line chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line3.png">
    LinePercentStacked,

    /// A Pie chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pie1.png">
    Pie,

    /// A Radar chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_radar1.png">
    Radar,

    /// A Radar with markers chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_radar2.png">
    RadarWithMarkers,

    /// A filled Radar chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_radar3.png">
    RadarFilled,

    /// A Scatter chart type. Scatter charts, and their variant, are the only
    /// type that have values (as opposed to categories) for their X-Axis. The
    /// default scatter chart in Excel has markers for each point but no
    /// connecting lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_scatter1.png">
    Scatter,

    /// A Scatter chart type where the points are connected by straight lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_scatter3.png">
    ScatterStraight,

    /// A Scatter chart type where the points have markers and are connected by
    /// straight lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_scatter2.png">
    ScatterStraightWithMarkers,

    /// A Scatter chart type where the points are connected by smoothed lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_scatter5.png">
    ScatterSmooth,

    /// A Scatter chart type where the points have markers and are connected by
    /// smoothed lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_scatter4.png">
    ScatterSmoothWithMarkers,
}

/// A struct to represent a Chart title.
#[derive(Clone)]
pub struct ChartTitle {
    pub(crate) range: ChartRange,
    pub(crate) cache_data: ChartSeriesCacheData,
    name: String,
    hidden: bool,
    is_horizontal: bool,
}

impl ChartTitle {
    pub(crate) fn new() -> ChartTitle {
        ChartTitle {
            range: ChartRange::new_from_range("", 0, 0, 0, 0),
            cache_data: ChartSeriesCacheData::new(),
            name: "".to_string(),
            hidden: false,
            is_horizontal: false,
        }
    }

    /// Add a title for a chart.
    ///
    /// Set the name (title) for the chart. The name is displayed above the
    /// chart.
    ///
    /// The name can be a simple string, a formula such as `Sheet1!$A$1` or a
    /// tuple with a sheet name, row and column such as `('Sheet1', 0, 0)`.
    ///
    /// # Arguments
    ///
    /// * `range` - The range property which can be one of the following generic
    ///   types:
    ///    - A simple string title.
    ///    - A string with an Excel like range formula such as `"Sheet1!$A$1"`.
    ///    - A tuple that can be used to create the range programmatically using
    ///      a sheet name and zero indexed row and column values like:
    ///      `("Sheet1", 0, 0)` (this gives the same range as the previous
    ///      string value).
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the chart title.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_title_set_name.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 50)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #
    /// #     // Create a simple Column chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Set the chart title.
    ///     chart.title().set_name("This is the chart title");
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_title_set_name.png">
    ///
    pub fn set_name<T>(&mut self, name: T) -> &mut ChartTitle
    where
        T: IntoChartRange,
    {
        self.range = name.new_chart_range();

        // If the name didn't convert to a populated range then it probably just
        // a simple string title.
        if !self.range.has_data() {
            self.name = self.range.range_string.clone();
        }

        self
    }

    /// Hide the chart title.
    ///
    /// By default Excel adds an automatic chart title to charts with a single
    /// series and a user defined series name. The none `set_hidden()` method
    /// turns this default title off.
    ///
    /// # Examples
    ///
    /// A simple chart example using the rust_xlsxwriter library.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_title_set_hidden.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 50)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #
    /// #     // Create a simple Column chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$3")
    ///         .set_name("Yearly results");
    ///
    ///     // Hide the default chart title.
    ///     chart.title().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file is shown below. Note how there is a default title of
    /// `"Yearly results"` in the `"=SERIES("Yearly
    /// results",,Sheet1!$A$1:$A$3,1)"` formula in Excel but that it isn't
    /// displayed on the chart.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_title_set_hidden.png">
    ///
    pub fn set_hidden(&mut self) -> &mut ChartTitle {
        self.hidden = true;
        self
    }
}
/// A struct to represent a Chart axis.
#[derive(Clone)]
pub struct ChartAxis {
    axis_type: ChartAxisType,
    axis_position: ChartAxisPosition,
    pub(crate) title: ChartTitle,
}

impl ChartAxis {
    pub(crate) fn new() -> ChartAxis {
        ChartAxis {
            axis_type: ChartAxisType::Value,
            axis_position: ChartAxisPosition::Bottom,
            title: ChartTitle::new(),
        }
    }

    /// Add a title for a chart axis.
    ///
    /// Set the name (title) for the chart axis.
    ///
    /// The name can be a simple string, a formula such as `Sheet1!$A$1` or a
    /// tuple with a sheet name, row and column such as `('Sheet1', 0, 0)`.
    ///
    /// # Arguments
    ///
    /// * `range` - The range property which can be one of the following generic
    ///   types:
    ///    - A simple string title.
    ///    - A string with an Excel like range formula such as `"Sheet1!$A$1"`.
    ///    - A tuple that can be used to create the range programmatically using
    ///      a sheet name and zero indexed row and column values like:
    ///      `("Sheet1", 0, 0)` (this gives the same range as the previous
    ///      string value).
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the title of chart axes.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_name.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 50)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #
    /// #     // Create a simple Column chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Set the chart axis titles.
    ///     chart.x_axis().set_name("Test number");
    ///     chart.y_axis().set_name("Sample length (mm)");
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_name.png">
    ///
    pub fn set_name<T>(&mut self, name: T) -> &mut ChartAxis
    where
        T: IntoChartRange,
    {
        self.title.set_name(name);
        self
    }
}

#[derive(Clone)]
pub(crate) enum ChartAxisType {
    Category,
    Value,
}

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
    use crate::{ChartRange, ChartSeries};
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let mut series1 = ChartSeries::new();
        series1
            .set_categories(("Sheet1", 0, 0, 4, 0))
            .set_values(("Sheet1", 0, 1, 4, 1))
            .set_category_cache(&["1", "2", "3", "4", "5"], true)
            .set_value_cache(&["2", "4", "6", "8", "10"], true);

        let mut series2 = ChartSeries::new();
        series2
            .set_categories("Sheet1!$A$1:$A$5")
            .set_values("Sheet1!$C$1:$C$5")
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

    #[test]
    fn test_range_from_string() {
        let range_string = "=Sheet1!$A$1:$A$5";
        let range = ChartRange::new_from_string(range_string);
        assert_eq!("Sheet1!$A$1:$A$5", range.formula());
        assert_eq!("Sheet1", range.sheet_name);

        let range_string = "Sheet1!$A$1:$A$5";
        let range = ChartRange::new_from_string(range_string);
        assert_eq!("Sheet1!$A$1:$A$5", range.formula());
        assert_eq!("Sheet1", range.sheet_name);

        let range_string = "Sheet 1!$A$1:$A$5";
        let range = ChartRange::new_from_string(range_string);
        assert_eq!("'Sheet 1'!$A$1:$A$5", range.formula());
        assert_eq!("Sheet 1", range.sheet_name);
    }
}
