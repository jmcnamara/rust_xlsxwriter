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
    ColNum, IntoColor, ObjectMovement, RowNum, XlsxColor, XlsxError, COL_MAX, ROW_MAX,
};

#[derive(Clone)]
/// The Chart struct is used to create an object to represent an chart that can
/// be inserted into a worksheet.
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_intro.png">
///
/// The `Chart` struct exposes other chart related structs that allow you to
/// configure the chart such as [`ChartSeries`], [`ChartAxis`] and
/// [`ChartTitle`].
///
/// Charts are added to the worksheets using the the
/// [`worksheet.insert_chart()`](crate::Worksheet::insert_chart) or
/// [`worksheet.insert_chart_with_offset()`](crate::Worksheet::insert_chart_with_offset)
/// methods.
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
///     // Create a new chart.
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
    pub(crate) legend: ChartLegend,
    pub(crate) chart_area_format: ChartFormat,
    pub(crate) plot_area_format: ChartFormat,
    grouping: ChartGrouping,
    default_cross_between: bool,
    default_num_format: String,
    has_overlap: bool,
    overlap: i8,
    gap: u16,
    style: u8,
    hole_size: u8,
    rotation: u16,
}

impl Chart {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new `Chart` struct.
    ///
    /// Create a new [`Chart`] object that can be configured and inserted into a
    /// worksheet using the
    /// [`worksheet.insert_chart()`][crate::Worksheet::insert_chart].
    ///
    /// Once you have create a chart you will need to add at least one data
    /// series via [`chart.add_series()`](Chart::add_series) and set a value range
    /// for that series using [`series.set_values()`][ChartSeries::set_values].
    /// See the example below.
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
    /// #     // Create a new chart.
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
            legend: ChartLegend::new(),
            chart_area_format: ChartFormat::new(),
            plot_area_format: ChartFormat::new(),
            grouping: ChartGrouping::Standard,
            default_cross_between: true,
            default_num_format: "General".to_string(),
            has_overlap: false,
            overlap: 0,
            gap: 150,
            style: 2,
            hole_size: 50,
            rotation: 0,
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

    /// Create and add a new chart series to a chart.
    ///
    /// Create and add a new chart series to a chart. The chart series
    /// represents the category and value ranges as well as formatting and
    /// display options. A chart in Excel must contain at least one data series.
    /// A series is represented by a [`ChartSeries`] struct.
    ///
    /// A chart series is usually created via this `add_series()` method.
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
    /// #     // Create a new chart.
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
        let mut series = ChartSeries::new();

        // The default Scatter chart has a hidden line with a standard width.
        if self.chart_type == ChartType::Scatter {
            series.set_format(
                ChartFormat::new().set_line(ChartLine::new().set_width(2.25).set_hidden()),
            );
        }

        // Turn off markers for chart types that can have markers but don't have
        // them by default.
        if self.chart_type == ChartType::ScatterStraight
            || self.chart_type == ChartType::ScatterSmooth
            || self.chart_group_type == ChartType::Line
            || self.chart_type == ChartType::Radar
        {
            series.marker = Some(ChartMarker::new().set_none().clone());
        }

        self.series.push(series);

        self.series.last_mut().unwrap()
    }

    /// Add a chart series to a chart.
    ///
    /// Add a standalone chart series to a chart. The chart series represents
    /// the category and value ranges as well as formatting and display options.
    /// A chart in Excel must contain at least one data series. A series is
    /// represented by a [`ChartSeries`] struct.
    ///
    /// A chart series is usually created via the
    /// [`chart.add_series()`](Chart::add_series) method, see above. However, if
    /// required you can create a standalone `ChartSeries` object and add it to
    /// a chart via this `chart.push_series()` method.
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
    /// #     // Create a new chart.
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
        let mut series = series.clone();

        // The default Scatter chart has a hidden line with a standard width.
        if self.chart_type == ChartType::Scatter {
            series.set_format(
                ChartFormat::new().set_line(ChartLine::new().set_width(2.25).set_hidden()),
            );
        }

        // Turn off markers for chart types that can have markers but don't have
        // them by default.
        if self.chart_type == ChartType::ScatterStraight
            || self.chart_type == ChartType::ScatterSmooth
            || self.chart_group_type == ChartType::Line
            || self.chart_type == ChartType::Radar
        {
            series.marker = Some(ChartMarker::new().set_none().clone());
        }

        self.series.push(series);

        self
    }

    /// Get the chart title object in order to set its properties.
    ///
    /// Get a reference to the chart's X-Axis [`ChartTitle`] object in order to
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
    /// #     // Create a new chart.
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
    /// Get a reference to the chart's X-Axis [`ChartAxis`] object in order to
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
    /// #     // Create a new chart.
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
    /// Get a reference to the chart's Y-Axis [`ChartAxis`] object in order to
    /// set its properties.
    ///
    /// See the [`chart.x_axis()`][Chart::x_axis] method above.
    ///
    pub fn y_axis(&mut self) -> &mut ChartAxis {
        &mut self.y_axis
    }

    /// Get the chart legend object in order to set its properties.
    ///
    /// Get a reference to the chart's [`ChartLegend`] object in order to set
    /// its properties.
    ///
    /// # Examples
    ///
    /// An example of getting the chart legend object and setting some of its
    /// properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_legend.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartLegendPosition, ChartType, Workbook, XlsxError};
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
    /// #     worksheet.write(1, 1, 35)?;
    /// #     worksheet.write(2, 1, 45)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///     chart.add_series().set_values("Sheet1!$B$1:$B$3");
    ///
    ///     // Turn on the chart legend and place it at the bottom of the chart.
    ///     chart.legend().set_position(ChartLegendPosition::Bottom);
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_legend.png">
    ///
    pub fn legend(&mut self) -> &mut ChartLegend {
        &mut self.legend
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
    /// An example showing all 48 default chart styles available in Excel 2007
    /// using rust_xlsxwriter.
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

    /// Set the formatting properties for the chart area.
    ///
    /// Set the formatting properties for a chart area via a [`ChartFormat`]
    /// object or a sub struct that implements [`IntoChartFormat`]. In Excel the
    /// chart area is the background area behind the chart.
    ///
    /// The formatting that can be applied via a [`ChartFormat`] object are:
    ///
    /// - `no_fill`: Turn of the fill for the chart object.
    /// - `solid_fill`: Set the [`ChartSolidFill`] properties.
    /// - `pattern_fill`: Set the [`ChartPatternFill`] properties.
    /// - `no_line`: Turn off the line/border for the chart object.
    /// - `line`: Set the [`ChartLine`] properties.
    ///
    /// # Arguments
    ///
    /// `format`: A [`ChartFormat`] struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// [`IntoChartFormat`] for details.
    ///
    /// # Examples
    ///
    /// An example of formatting the chart "area" of a chart. In Excel the chart
    /// area is the background area behind the chart.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_chart_area_format.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFormat, ChartSolidFill, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///         chart.set_chart_area_format(
    ///             ChartFormat::new().set_solid_fill(
    ///                 ChartSolidFill::new()
    ///                     .set_color("#FFFFB3")
    ///             ),
    ///         );
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
    /// src="https://rustxlsxwriter.github.io/images/chart_set_chart_area_format.png">
    ///
    pub fn set_chart_area_format<T>(&mut self, format: T) -> &mut Chart
    where
        T: IntoChartFormat,
    {
        self.chart_area_format = format.new_chart_format();
        self
    }

    /// Set the formatting properties for the plot area.
    ///
    /// Set the formatting properties for a chart plot area via a
    /// [`ChartFormat`] object. In Excel the plot area is the area between the
    /// axes on which the chart series are plotted.
    ///
    /// The formatting that can be applied via a [`ChartFormat`] object are:
    ///
    /// - `no_fill`: Turn of the fill for the chart object.
    /// - `solid_fill`: Set the [`ChartSolidFill`] properties.
    /// - `pattern_fill`: Set the [`ChartPatternFill`] properties.
    /// - `no_line`: Turn off the line/border for the chart object.
    /// - `line`: Set the [`ChartLine`] properties.
    ///
    /// # Arguments
    ///
    /// `format`: A [`ChartFormat`] struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// [`IntoChartFormat`] for details.
    ///
    /// # Examples
    ///
    /// An example of formatting the chart "area" of a chart. In Excel the plot
    /// area is the area between the axes on which the chart series are plotted.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_plot_area_format.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFormat, ChartSolidFill, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///         chart.set_plot_area_format(
    ///             ChartFormat::new().set_solid_fill(
    ///                 ChartSolidFill::new()
    ///                     .set_color("#FFFFB3")
    ///             ),
    ///         );
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
    /// src="https://rustxlsxwriter.github.io/images/chart_set_plot_area_format.png">
    ///
    pub fn set_plot_area_format<T>(&mut self, format: T) -> &mut Chart
    where
        T: IntoChartFormat,
    {
        self.plot_area_format = format.new_chart_format();
        self
    }

    /// Set the Pie/Doughnut chart rotation.
    ///
    /// The `set_rotation()` method is used to set the rotation of the first
    /// segment of a Pie/Doughnut chart. This has the effect of rotating the
    /// entire chart.
    ///
    /// # Arguments
    ///
    /// * `rotation`: The rotation of the first segment of a Pie/Doughnut chart.
    /// The range is 0 <= rotation <= 360 and the default is 0.
    ///
    ///
    /// # Examples
    ///
    /// An example of formatting the chart rotation for pie and doughnut charts.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_rotation.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Pie);
    ///
    ///     // Add a data series with formatting.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Set the rotation of the chart.
    ///     chart.set_rotation(270);
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_set_rotation.png">
    ///
    pub fn set_rotation(&mut self, rotation: u16) -> &mut Chart {
        if (0..=360).contains(&rotation) {
            self.rotation = rotation;
        }
        self
    }

    /// Set the hole size for a Doughnut chart.
    ///
    /// Set the center hole size for a Doughnut chart.
    ///
    /// # Arguments
    ///
    /// * `hole_size`: The hole size for a Doughnut chart. The range is 0 <=
    /// `hole_size` <= 90 and the default is 50.
    ///
    ///
    /// # Examples
    ///
    /// An example of formatting the chart hole size for doughnut charts.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_hole_size.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Doughnut);
    ///
    ///     // Add a data series with formatting.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Set the home size of the chart.
    ///     chart.set_hole_size(80);
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_set_hole_size.png">
    ///
    pub fn set_hole_size(&mut self, hole_size: u8) -> &mut Chart {
        if (0..=90).contains(&hole_size) {
            self.hole_size = hole_size;
        }
        self
    }

    /// Set the width of the chart.
    ///
    /// The default width of an Excel chart is 480 pixels. The `set_width()`
    /// method allows you to set it to some other non-zero size.
    ///
    /// # Arguments
    ///
    /// * `width` - The chart width in pixels.
    ///
    /// # Examples
    ///
    /// A simple chart example using the rust_xlsxwriter library.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_width.rs
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
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Hide the legend, for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Resize the chart.
    ///     chart.set_height(200).set_width(240);
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_set_width.png">
    ///
    pub fn set_width(&mut self, width: u32) -> &mut Chart {
        if width == 0 {
            return self;
        }

        self.width = width as f64;
        self
    }

    /// Set the height of the chart.
    ///
    /// The default height of an Excel chart is 480 pixels. The `set_height()`
    /// method allows you to set it to some other non-zero size. See the example
    /// above.
    ///
    /// # Arguments
    ///
    /// * `height` - The chart height in pixels.
    ///
    pub fn set_height(&mut self, height: u32) -> &mut Chart {
        if height == 0 {
            return self;
        }

        self.height = height as f64;
        self
    }

    /// Set the height scale for the chart.
    ///
    /// Set the height scale for the chart relative to 1.0/100%. This is a
    /// syntactic alternative to [`chart.set_height()`](Chart::set_height).
    ///
    /// # Arguments
    ///
    /// * `scale` - The scale ratio.
    ///
    pub fn set_scale_height(&mut self, scale: f64) -> &mut Chart {
        if scale <= 0.0 {
            return self;
        }

        self.scale_height = scale;
        self
    }

    /// Set the width scale for the chart.
    ///
    /// Set the width scale for the chart relative to 1.0/100%. This is a
    /// syntactic alternative to [`chart.set_width()`](Chart::set_width).
    ///
    /// # Arguments
    ///
    /// * `scale` - The scale ratio.
    ///
    pub fn set_scale_width(&mut self, scale: f64) -> &mut Chart {
        if scale <= 0.0 {
            return self;
        }

        self.scale_width = scale;
        self
    }

    /// Set the alt text for the chart.
    ///
    /// Set the alt text for the chart to help accessibility. The alt text is
    /// used with screen readers to help people with visual disabilities.
    ///
    /// See the following Microsoft documentation on [Everything you need to
    /// know to write effective alt
    /// text](https://support.microsoft.com/en-us/office/everything-you-need-to-know-to-write-effective-alt-text-df98f884-ca3d-456c-807b-1a1fa82f5dc2).
    ///
    /// # Arguments
    ///
    /// * `alt_text` - The alt text string to add to the chart.
    ///
    pub fn set_alt_text(&mut self, alt_text: &str) -> &mut Chart {
        self.alt_text = alt_text.to_string();
        self
    }

    /// Mark a chart as decorative.
    ///
    /// Charts don't always need an alt text description. Some charts may contain
    /// little or no useful visual information. Such charts can be marked as
    /// "decorative" so that screen readers can inform the users that they don't
    /// contain important information.
    ///
    /// # Arguments
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn set_decorative(&mut self, enable: bool) -> &mut Chart {
        self.decorative = enable;
        self
    }

    /// Set the object movement options for a chart.
    ///
    /// Set the option to define how an chart will behave in Excel if the cells
    /// under the chart are moved, deleted, or have their size changed. In Excel
    /// the options are:
    ///
    /// 1. Move and size with cells. Default for charts.
    /// 2. Move but don't size with cells.
    /// 3. Don't move or size with cells.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/object_movement.png">
    ///
    /// These values are defined in the [`ObjectMovement`] enum.
    ///
    pub fn set_object_movement(&mut self, option: ObjectMovement) -> &mut Chart {
        self.object_movement = option;
        self
    }

    /// Check a chart instance for configuration errors.
    ///
    /// Charts are validated using this methods when they are added to a
    /// worksheet using the
    /// [`worksheet.insert_chart()`](crate::Worksheet::insert_chart) or
    /// [`worksheet.insert_chart_with_offset()`](crate::Worksheet::insert_chart_with_offset)
    /// methods. However, you can also call `chart.validate()` directly.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::ChartError`] - A general error that is raised when a
    /// chart parameter is incorrect or a chart is configured incorrectly.
    ///
    pub fn validate(&mut self) -> Result<&mut Chart, XlsxError> {
        // Check for chart without series.
        if self.series.is_empty() {
            return Err(XlsxError::ChartError(
                "Chart must contain at least one series".to_string(),
            ));
        }

        for series in self.series.iter() {
            // Check for a series without a values range.
            if !series.value_range.has_data() {
                return Err(XlsxError::ChartError(
                    "Chart series must contain a values range".to_string(),
                ));
            }

            // Check for scatter charts without category ranges. It is optional
            // for all other types.
            if self.chart_group_type == ChartType::Scatter && !series.category_range.has_data() {
                return Err(XlsxError::ChartError(
                    "Scatter style charts must contain a categories range".to_string(),
                ));
            }

            // Validate the series values range.
            series.value_range.validate()?;

            // Validate the series category range.
            if series.category_range.has_data() {
                series.category_range.validate()?;
            }
        }

        Ok(self)
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

        if self.gap != 150 {
            // Write the c:gapWidth element.
            self.write_gap_width(self.gap);
        }

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

        if self.gap != 150 {
            // Write the c:gapWidth element.
            self.write_gap_width(self.gap);
        }

        if self.overlap != 0 {
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

        // Write the c:spPr element.
        self.write_sp_pr(&self.chart_area_format.clone());

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
        } else if title.format.has_formatting() {
            self.write_title_format_only(title);
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
                self.write_area_chart();
            }

            ChartType::Bar | ChartType::BarStacked | ChartType::BarPercentStacked => {
                self.write_bar_chart();
            }

            ChartType::Column | ChartType::ColumnStacked | ChartType::ColumnPercentStacked => {
                self.write_column_chart();
            }

            ChartType::Doughnut => self.write_doughnut_chart(),

            ChartType::Line | ChartType::LineStacked | ChartType::LinePercentStacked => {
                self.write_line_chart();
            }

            ChartType::Pie => self.write_pie_chart(),

            ChartType::Radar | ChartType::RadarWithMarkers | ChartType::RadarFilled => {
                self.write_radar_chart();
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

        // Write the c:spPr element.
        self.write_sp_pr(&self.plot_area_format.clone());

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
            attributes.push(("val", "smoothMarker".to_string()));
        } else {
            attributes.push(("val", "lineMarker".to_string()));
        }

        self.writer
            .xml_empty_tag_attr("c:scatterStyle", &attributes);
    }

    // Write the <c:ser> element.
    fn write_series(&mut self) {
        for (index, series) in self.series.clone().iter_mut().enumerate() {
            self.writer.xml_start_tag("c:ser");

            // Copy a series overlap to the parent chart.
            if series.overlap != 0 {
                self.overlap = series.overlap;
            }

            // Copy a series gap to the parent chart.
            if series.gap != 150 {
                self.gap = series.gap;
            }

            // Write the c:idx element.
            self.write_idx(index);

            // Write the c:order element.
            self.write_order(index);

            self.write_series_title(&series.title);

            // Write the c:spPr element.
            self.write_sp_pr(&series.format);

            if let Some(marker) = &series.marker {
                if !marker.automatic {
                    // Write the c:marker element.
                    self.write_marker(marker);
                }
            }

            // Write the point formatting for the series.
            if !series.points.is_empty() {
                self.write_d_pt(&series.points);
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
        for (index, series) in self.series.clone().iter_mut().enumerate() {
            self.writer.xml_start_tag("c:ser");

            // Write the c:idx element.
            self.write_idx(index);

            // Write the c:order element.
            self.write_order(index);

            self.write_series_title(&series.title);

            if let Some(marker) = &series.marker {
                if !marker.automatic {
                    // Write the c:marker element.
                    self.write_marker(marker);
                }
            }

            // Add default scatter line formatting to the series data unless it
            // has already been specified by the user.
            if self.chart_type == ChartType::Scatter && series.format.line.is_none() {
                let mut line = ChartLine::new();
                line.set_width(2.25);
                series.format.line = Some(line);
            }

            // Write the c:spPr formatting element.
            self.write_sp_pr(&series.format);

            // Write the point formatting for the series.
            if !series.points.is_empty() {
                self.write_d_pt(&series.points);
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

    // Write the <c:dPt> element.
    fn write_d_pt(&mut self, points: &[ChartPoint]) {
        let has_marker =
            self.chart_group_type == ChartType::Scatter || self.chart_group_type == ChartType::Line;

        // Write the point formatting for the series.
        for (index, point) in points.iter().enumerate() {
            if point.is_not_default() {
                self.writer.xml_start_tag("c:dPt");
                self.write_idx(index);

                if has_marker {
                    self.writer.xml_start_tag("c:marker");
                }

                if point.format.has_formatting() {
                    // Write the c:spPr formatting element.
                    self.write_sp_pr(&point.format);
                }

                if has_marker {
                    self.writer.xml_end_tag("c:marker");
                }

                self.writer.xml_end_tag("c:dPt");
            }
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
            if !value.is_empty() {
                self.write_pt(index, value);
            }
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

        if self.x_axis.format.has_formatting() {
            // Write the c:spPr formatting element.
            self.write_sp_pr(&self.x_axis.format.clone());
        }

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

        if self.y_axis.format.has_formatting() {
            // Write the c:spPr formatting element.
            self.write_sp_pr(&self.y_axis.format.clone());
        }

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

        if self.x_axis.format.has_formatting() {
            // Write the c:spPr formatting element.
            self.write_sp_pr(&self.x_axis.format.clone());
        }

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
        if self.legend.hidden {
            return;
        }

        self.writer.xml_start_tag("c:legend");

        // Write the c:legendPos element.
        self.write_legend_pos();

        // Write the c:layout element.
        self.write_layout();

        // Write the c:spPr formatting element.
        self.write_sp_pr(&self.legend.format.clone());

        // Write the c:overlay element.
        self.write_overlay();

        if self.chart_type == ChartType::Pie || self.chart_type == ChartType::Doughnut {
            // Write the c:txPr element.
            self.write_tx_pr_pie();
        }

        self.writer.xml_end_tag("c:legend");
    }

    // Write the <c:legendPos> element.
    fn write_legend_pos(&mut self) {
        let attributes = vec![("val", self.legend.position.to_string())];

        self.writer.xml_empty_tag_attr("c:legendPos", &attributes);
    }

    // Write the <c:overlay> element.
    fn write_overlay(&mut self) {
        if !self.legend.has_overlay {
            return;
        }

        let attributes = vec![("val", "1".to_string())];

        self.writer.xml_empty_tag_attr("c:overlay", &attributes);
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
    fn write_marker(&mut self, marker: &ChartMarker) {
        self.writer.xml_start_tag("c:marker");

        // Write the c:symbol element.
        self.write_symbol(marker);

        if marker.size != 0 {
            // Write the c:size element.
            self.write_size(marker.size);
        }

        if marker.format.has_formatting() {
            // Write the c:spPr formatting element.
            self.write_sp_pr(&marker.format);
        }

        self.writer.xml_end_tag("c:marker");
    }

    // Write the <c:symbol> element.
    fn write_symbol(&mut self, marker: &ChartMarker) {
        let mut attributes = vec![];

        if let Some(marker_type) = marker.marker_type {
            attributes.push(("val", marker_type.to_string()));
        } else if marker.none {
            attributes.push(("val", "none".to_string()));
        }

        self.writer.xml_empty_tag_attr("c:symbol", &attributes);
    }

    // Write the <c:size> element.
    fn write_size(&mut self, size: u8) {
        let attributes = vec![("val", size.to_string())];

        self.writer.xml_empty_tag_attr("c:size", &attributes);
    }

    // Write the <c:varyColors> element.
    fn write_vary_colors(&mut self) {
        let attributes = vec![("val", "1".to_string())];

        self.writer.xml_empty_tag_attr("c:varyColors", &attributes);
    }

    // Write the <c:firstSliceAng> element.
    fn write_first_slice_ang(&mut self) {
        let attributes = vec![("val", self.rotation.to_string())];

        self.writer
            .xml_empty_tag_attr("c:firstSliceAng", &attributes);
    }

    // Write the <c:holeSize> element.
    fn write_hole_size(&mut self) {
        let attributes = vec![("val", self.hole_size.to_string())];

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
    fn write_sp_pr(&mut self, format: &ChartFormat) {
        if !format.has_formatting() {
            return;
        }

        self.writer.xml_start_tag("c:spPr");

        if format.no_fill {
            self.writer.xml_empty_tag("a:noFill");
        } else if let Some(solid_fill) = &format.solid_fill {
            // Write the a:solidFill element.
            self.write_a_solid_fill(solid_fill.color, solid_fill.transparency);
        } else if let Some(pattern_fill) = &format.pattern_fill {
            // Write the a:pattFill element.
            self.write_a_patt_fill(pattern_fill);
        }

        if format.no_line {
            // Write a default line with no fill.
            self.write_a_ln_none();
        } else if let Some(line) = &format.line {
            // Write the a:ln element.
            self.write_a_ln(line);
        }

        self.writer.xml_end_tag("c:spPr");
    }

    // Write the <a:ln> element.
    fn write_a_ln(&mut self, line: &ChartLine) {
        let mut attributes = vec![];

        if let Some(width) = &line.width {
            /* Round width to nearest 0.25, like Excel. */
            let width = ((*width + 0.125) * 4.0).floor() / 4.0;

            /* Convert to Excel internal units. */
            let width = (12700.0 * width).ceil() as u32;

            attributes.push(("w", width.to_string()));
        }

        if line.color != XlsxColor::Default
            || line.dash_type != ChartLineDashType::Solid
            || line.hidden
        {
            self.writer.xml_start_tag_attr("a:ln", &attributes);

            if line.hidden {
                // Write the a:noFill element.
                self.write_a_no_fill();
            } else {
                if line.color != XlsxColor::Default {
                    // Write the a:solidFill element.
                    self.write_a_solid_fill(line.color, line.transparency);
                }

                if line.dash_type != ChartLineDashType::Solid {
                    // Write the a:prstDash element.
                    self.write_a_prst_dash(line);
                }
            }

            self.writer.xml_end_tag("a:ln");
        } else {
            self.writer.xml_empty_tag_attr("a:ln", &attributes);
        }
    }

    // Write the <a:ln> element.
    fn write_a_ln_none(&mut self) {
        self.writer.xml_start_tag("a:ln");

        // Write the a:noFill element.
        self.write_a_no_fill();

        self.writer.xml_end_tag("a:ln");
    }

    // Write the <a:solidFill> element.
    fn write_a_solid_fill(&mut self, color: XlsxColor, transparency: u8) {
        self.writer.xml_start_tag("a:solidFill");

        // Write the color element.
        self.write_color(color, transparency);

        self.writer.xml_end_tag("a:solidFill");
    }

    // Write the <a:pattFill> element.
    fn write_a_patt_fill(&mut self, fill: &ChartPatternFill) {
        let attributes = vec![("prst", fill.pattern.to_string())];

        self.writer.xml_start_tag_attr("a:pattFill", &attributes);

        if fill.foreground_color != XlsxColor::Default {
            // Write the <a:fgClr> element.
            self.writer.xml_start_tag("a:fgClr");
            self.write_color(fill.foreground_color, 0);
            self.writer.xml_end_tag("a:fgClr");
        }

        if fill.background_color != XlsxColor::Default {
            // Write the <a:bgClr> element.
            self.writer.xml_start_tag("a:bgClr");
            self.write_color(fill.background_color, 0);
            self.writer.xml_end_tag("a:bgClr");
        }

        self.writer.xml_end_tag("a:pattFill");
    }

    // Write the <a:srgbClr> element.
    fn write_color(&mut self, color: XlsxColor, transparency: u8) {
        match color {
            XlsxColor::Theme(_, _) => {
                let (scheme, lum_mod, lum_off) = color.chart_scheme();
                if !scheme.is_empty() {
                    // Write the a:schemeClr element.
                    self.write_a_scheme_clr(scheme, lum_mod, lum_off, transparency);
                }
            }
            _ => {
                let attributes = vec![("val", color.rgb_hex_value())];

                if transparency > 0 {
                    self.writer.xml_start_tag_attr("a:srgbClr", &attributes);

                    // Write the a:alpha element.
                    self.write_a_alpha(transparency);

                    self.writer.xml_end_tag("a:srgbClr");
                } else {
                    self.writer.xml_empty_tag_attr("a:srgbClr", &attributes);
                }
            }
        }
    }

    // Write the <a:schemeClr> element.
    fn write_a_scheme_clr(&mut self, scheme: String, lum_mod: u32, lum_off: u32, transparency: u8) {
        let attributes = vec![("val", scheme)];

        if lum_mod > 0 || lum_off > 0 || transparency > 0 {
            self.writer.xml_start_tag_attr("a:schemeClr", &attributes);

            if lum_mod > 0 {
                // Write the a:lumMod element.
                self.write_a_lum_mod(lum_mod);
            }

            if lum_off > 0 {
                // Write the a:lumOff element.
                self.write_a_lum_off(lum_off);
            }

            if transparency > 0 {
                // Write the a:alpha element.
                self.write_a_alpha(transparency);
            }

            self.writer.xml_end_tag("a:schemeClr");
        } else {
            self.writer.xml_empty_tag_attr("a:schemeClr", &attributes);
        }
    }

    // Write the <a:lumMod> element.
    fn write_a_lum_mod(&mut self, lum_mod: u32) {
        let attributes = vec![("val", lum_mod.to_string())];

        self.writer.xml_empty_tag_attr("a:lumMod", &attributes);
    }

    // Write the <a:lumOff> element.
    fn write_a_lum_off(&mut self, lum_off: u32) {
        let attributes = vec![("val", lum_off.to_string())];

        self.writer.xml_empty_tag_attr("a:lumOff", &attributes);
    }

    // Write the <a:alpha> element.
    fn write_a_alpha(&mut self, transparency: u8) {
        let transparency = (100 - transparency) as u32 * 1000;

        let attributes = vec![("val", transparency.to_string())];

        self.writer.xml_empty_tag_attr("a:alpha", &attributes);
    }

    // Write the <a:noFill> element.
    fn write_a_no_fill(&mut self) {
        self.writer.xml_empty_tag("a:noFill");
    }

    // Write the <a:prstDash> element.
    fn write_a_prst_dash(&mut self, line: &ChartLine) {
        let attributes = vec![("val", line.dash_type.to_string())];

        self.writer.xml_empty_tag_attr("a:prstDash", &attributes);
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

    // Write the <c:gapWidth> element.
    fn write_gap_width(&mut self, gap: u16) {
        let attributes = vec![("val", gap.to_string())];

        self.writer.xml_empty_tag_attr("c:gapWidth", &attributes);
    }

    // Write the <c:overlap> element.
    fn write_overlap(&mut self) {
        let attributes = vec![("val", self.overlap.to_string())];

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

        if title.format.has_formatting() {
            // Write the c:spPr formatting element.
            self.write_sp_pr(&title.format.clone());
        } else {
            // Write the c:txPr element.
            self.write_tx_pr(title.is_horizontal);
        }

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

        if title.format.has_formatting() {
            // Write the c:spPr element.
            self.write_sp_pr(&title.format.clone());
        }

        self.writer.xml_end_tag("c:title");
    }

    // Write the <c:title> element.
    fn write_title_format_only(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag("c:title");

        // Write the c:layout element.
        self.write_layout();

        // Write the c:spPr element.
        self.write_sp_pr(&title.format.clone());

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

/// Trait to map types into a `ChartFormat`.
///
/// The `IntoChartFormat` trait provides a syntactic shortcut for the
/// `chart.*.set_format()` methods that take [`ChartFormat`] as a parameter.
///
/// The [`ChartFormat`] struct mirrors the Excel Chart element formatting dialog
/// and has several sub-structs such as:
///
/// - [`ChartLine`]
/// - [`ChartSolidFill`]
/// - [`ChartPatternFill`]
///
/// In order to pass one of these sub-structs as a parameter you would normally
/// have to create a [`ChartFormat`] first and then add the sub-struct, as shown
/// in the first part of the example below. However, since this is a little
/// verbose if you just want to format one of the sub-properties the
/// `IntoChartFormat` trait will accept the sub-structs listed above and create
/// a parent [`ChartFormat`] instance to wrap it in, see the second part of the
/// example below.
///
/// # Examples
///
/// An example of passing chart formatting parameters via the
/// [`IntoChartFormat`] trait
///
/// ```
/// # // This code is available in examples/doc_into_chart_format.rs
/// #
/// # use rust_xlsxwriter::{Chart, ChartFormat, ChartSolidFill, ChartType, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the chart.
/// #     worksheet.write(0, 0, 10)?;
/// #     worksheet.write(1, 0, 40)?;
/// #     worksheet.write(2, 0, 50)?;
/// #     worksheet.write(0, 1, 20)?;
/// #     worksheet.write(1, 1, 10)?;
/// #     worksheet.write(2, 1, 50)?;
/// #
/// #     // Create a new chart.
///     let mut chart = Chart::new(ChartType::Column);
///
///     // Add formatting via ChartFormat and a ChartSolidFill sub struct.
///     chart
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$3")
///         .set_format(ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#40EABB")));
///
///     // Add formatting using a ChartSolidFill struct directly.
///     chart
///         .add_series()
///         .set_values("Sheet1!$B$1:$B$3")
///         .set_format(ChartSolidFill::new().set_color("#AAC3F2"));
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
/// <img src="https://rustxlsxwriter.github.io/images/into_chart_format.png">
///
///
/// ```
/// # // This code is available in examples/doc_into_chart_format.rs
/// #
/// # use rust_xlsxwriter::{Chart, ChartFormat, ChartSolidFill, ChartType, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the chart.
/// #     worksheet.write(0, 0, 10)?;
/// #     worksheet.write(1, 0, 40)?;
/// #     worksheet.write(2, 0, 50)?;
/// #     worksheet.write(0, 1, 20)?;
/// #     worksheet.write(1, 1, 10)?;
/// #     worksheet.write(2, 1, 50)?;
/// #
/// #     // Create a new chart.
///     let mut chart = Chart::new(ChartType::Column);
///
///     // Add formatting via ChartFormat and a ChartSolidFill sub struct.
///     chart
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$3")
///         .set_format(ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#40EABB")));
///
///     // Add formatting using a ChartSolidFill struct directly.
///     chart
///         .add_series()
///         .set_values("Sheet1!$B$1:$B$3")
///         .set_format(ChartSolidFill::new().set_color("#AAC3F2"));
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
/// <img src="https://rustxlsxwriter.github.io/images/into_chart_format.png">
///
///
pub trait IntoChartFormat {
    /// Trait function to turn a type into [`ChartFormat`].
    fn new_chart_format(&self) -> ChartFormat;
}

impl IntoChartFormat for &mut ChartFormat {
    fn new_chart_format(&self) -> ChartFormat {
        (*self).clone()
    }
}

impl IntoChartFormat for &mut ChartLine {
    fn new_chart_format(&self) -> ChartFormat {
        ChartFormat::new().set_line(self).clone()
    }
}

impl IntoChartFormat for &mut ChartSolidFill {
    fn new_chart_format(&self) -> ChartFormat {
        ChartFormat::new().set_solid_fill(self).clone()
    }
}

impl IntoChartFormat for &mut ChartPatternFill {
    fn new_chart_format(&self) -> ChartFormat {
        ChartFormat::new().set_pattern_fill(self).clone()
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
///     // Create a new chart.
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
    pub(crate) format: ChartFormat,
    pub(crate) marker: Option<ChartMarker>,
    pub(crate) points: Vec<ChartPoint>,
    pub(crate) gap: u16,
    pub(crate) overlap: i8,
}

#[allow(clippy::new_without_default)]
impl ChartSeries {
    /// Create a new chart series object.
    ///
    /// Create a new chart series object. A chart in Excel must contain at least
    /// one data series. The `ChartSeries` struct represents the category and
    /// value ranges, and the formatting and options for the chart series.
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
    /// #     // Create a new chart.
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
    /// An example of creating a chart series as a standalone object and then
    /// adding it to a chart via the [`chart.push_series()`](Chart::add_series)
    /// method.
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
    /// #     // Create a new chart.
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
            format: ChartFormat::new(),
            marker: None,
            points: vec![],
            gap: 150,
            overlap: 0,
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
    /// #     // Create a new chart.
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
    /// #     // Create a new chart.
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
    /// #     // Create a new chart.
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

    /// Set the formatting properties for a chart series.
    ///
    /// Set the formatting properties for a chart series via a [`ChartFormat`]
    /// object or a sub struct that implements [`IntoChartFormat`].
    ///
    /// The formatting that can be applied via a [`ChartFormat`] object are:
    ///
    /// - `no_fill`: Turn of the fill for the chart object.
    /// - `solid_fill`: Set the [`ChartSolidFill`] properties.
    /// - `pattern_fill`: Set the [`ChartPatternFill`] properties.
    /// - `no_line`: Turn off the line/border for the chart object.
    /// - `line`: Set the [`ChartLine`] properties.
    ///
    /// # Arguments
    ///
    /// `format`: A [`ChartFormat`] struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// [`IntoChartFormat`] for details.
    ///
    pub fn set_format<T>(&mut self, format: T) -> &mut ChartSeries
    where
        T: IntoChartFormat,
    {
        self.format = format.new_chart_format();
        self
    }

    /// Set the markers for a chart series.
    ///
    /// Set the markers and marker properties for a data series using a
    /// [`ChartMarker`] instance. In general only Line, Scatter and Radar chart
    /// support markers.
    ///
    /// # Examples
    ///
    /// An example of adding markers to a line chart.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_marker.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartFormat, ChartMarker, ChartMarkerType, ChartSolidFill, ChartType, Workbook,
    /// #     XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_marker(
    ///             ChartMarker::new()
    ///                 .set_type(ChartMarkerType::Square)
    ///                 .set_size(10)
    ///                 .set_format(
    ///                     ChartFormat::new().set_solid_fill(
    ///                         ChartSolidFill::new().set_color("#FF0000")),
    ///                 ),
    ///         );
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker.png">
    ///
    pub fn set_marker(&mut self, marker: &ChartMarker) -> &mut ChartSeries {
        self.marker = Some(marker.clone());
        self
    }

    /// Set the formatting and properties for points in a chart series.
    ///
    /// The meaning of "point" varies between chart types. For a Line chart a point
    /// is a line segment; in a Column chart a point is a an individual bar; and in
    /// a Pie chart a point is a pie segment.
    ///
    /// A point is represented by the [`ChartPoint`] struct.
    ///
    /// Chart points are most commonly used for Pie and Doughnut charts to format
    /// individual segments of the chart. In all other chart types the formatting
    /// happens at the chart series level.
    ///
    ///
    /// # Arguments
    ///
    /// `points`: A slice of [`ChartPoint`] objects.
    ///
    /// # Examples
    ///
    /// An example of formatting the individual segments of a Pie chart.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_points.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartFormat, ChartPoint, ChartSolidFill, ChartType, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 15)?;
    /// #     worksheet.write(1, 0, 15)?;
    /// #     worksheet.write(2, 0, 30)?;
    /// #
    ///     // Some point object with formatting to use in the Pie chart.
    ///     let points = vec![
    ///         ChartPoint::new().set_format(
    ///             ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FF0000")),
    ///         ),
    ///         ChartPoint::new().set_format(
    ///             ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FFC000")),
    ///         ),
    ///         ChartPoint::new().set_format(
    ///             ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FFFF00")),
    ///         ),
    ///     ];
    ///
    ///     // Create a simple Pie chart.
    ///     let mut chart = Chart::new(ChartType::Pie);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$3")
    ///         .set_points(&points);
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_set_points.png">
    ///
    pub fn set_points(&mut self, points: &[ChartPoint]) -> &mut ChartSeries {
        self.points = points.to_vec();
        self
    }

    /// Set the colors for points in a chart series.
    ///
    /// As explained above, in the section on
    /// [`set_points`](ChartSeries::set_points), the most common use case for
    /// point formatting is to set the formatting of individual segments of Pie
    /// charts, or in particular to set the colors of pie segments. For this
    /// simple use case the [`set_points`](ChartSeries::set_points) method can be
    /// overly verbose.
    ///
    /// As a syntactic shortcut the `set_point_colors()` method allow you to set
    /// the colors of chart points with a simpler interface.
    ///
    /// Compare the example below with the previous more general example which
    /// both produce the same result.
    ///
    /// # Arguments
    ///
    /// `colors`: a slice of [`XlsxColor`] enum values or types that will
    /// convert into [`XlsxColor`] via [`IntoColor`].
    ///
    ///
    ///
    /// # Examples
    ///
    /// An example of setting the individual segment colors of a Pie chart.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_point_colors.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 15)?;
    /// #     worksheet.write(1, 0, 15)?;
    /// #     worksheet.write(2, 0, 30)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Pie);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$3")
    ///         .set_point_colors(&["#FF000", "#FFC000", "#FFFF00"]);
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_set_points.png">
    ///
    pub fn set_point_colors<T>(&mut self, colors: &[T]) -> &mut ChartSeries
    where
        T: IntoColor + Copy,
    {
        self.points = colors
            .iter()
            .map(|color| ChartPoint::new().set_format(ChartSolidFill::new().set_color(*color)))
            .collect();
        self
    }

    /// Set the series overlap for a chart/bar chart.
    ///
    /// Set the overlap between series in a Bar/Column chart. The range is -100
    /// <= overlap <= 100 and the default is 0.
    ///
    /// Note, In Excel this property is only available for Bar and Column charts
    /// and also only needs to be applied to one of the data series of the
    /// chart.
    ///
    /// # Arguments
    ///
    /// * `overlap`: Overlap percentage of columns in Bar/Column charts. The
    /// range is -100 <= overlap <= 100 and the default is 0.
    ///
    /// # Examples
    ///
    /// A example of setting the chart series gap and overlap. Note that it only
    /// needs to be applied to one of the series in the chart.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_series_set_overlap.rs
    /// #
    /// # use rust_xlsxwriter::*;
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add the worksheet data that the charts will refer to.
    /// #     let data = [[105, 150, 130, 90], [50, 120, 100, 110]];
    /// #     for (col_num, col_data) in data.iter().enumerate() {
    /// #         for (row_num, row_data) in col_data.iter().enumerate() {
    /// #             worksheet.write(row_num as u32, col_num as u16, *row_data)?;
    /// #         }
    /// #     }
    /// #
    /// #     // Create a new column chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Configure the data series and add a gap/overlap. Note that it only needs
    ///     // to be applied to one of the series in the chart.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$4")
    ///         .set_overlap(37)
    ///         .set_gap(70);
    ///
    ///     chart.add_series().set_values("Sheet1!$B$1:$B$4");
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(1, 3, &chart)?;
    ///
    ///     workbook.save("chart.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_series_set_overlap.png">
    ///
    pub fn set_overlap(&mut self, overlap: i8) -> &mut ChartSeries {
        if (-100..=100).contains(&overlap) {
            self.overlap = overlap;
        }
        self
    }

    /// Set the gap width for a chart/bar chart.
    ///
    /// Set the gap width between series in a Bar/Column chart. The range is 0
    /// <= gap <= 500 and the default is 150.
    ///
    /// Note, In Excel this property is only available for Bar and Column charts
    /// and also only needs to be applied to one of the data series of the
    /// chart.
    ///
    /// # Arguments
    ///
    /// * `gap`: Gap percentage of columns in Bar/Column charts. The range is 0
    /// <= gap <= 500 and the default is 150.
    ///
    /// See the example for [`series.set_overlap()`](ChartSeries::set_overlap)
    /// above.
    ///
    pub fn set_gap(&mut self, gap: u16) -> &mut ChartSeries {
        if gap <= 500 {
            self.gap = gap;
        }
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
            data: data.iter().map(std::string::ToString::to_string).collect(),
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
            data: data.iter().map(|s| (*s).to_string()).collect(),
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
    // Create a new range from a sheet 5 tuple.
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

    // Create a new range from an Excel range formula.
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

    // Convert the row/col range into a chart range string.
    pub(crate) fn formula(&self) -> String {
        utility::chart_range_abs(
            &self.sheet_name,
            self.first_row,
            self.first_col,
            self.last_row,
            self.last_col,
        )
    }

    // Unique key to identify/find the range of values to build the cache.
    pub(crate) fn key(&self) -> (String, RowNum, ColNum, RowNum, ColNum) {
        (
            self.sheet_name.clone(),
            self.first_row,
            self.first_col,
            self.last_row,
            self.last_col,
        )
    }

    // Check that the range has data.
    pub(crate) fn has_data(&self) -> bool {
        !self.sheet_name.is_empty()
    }

    // Check that the row/column values in the range are valid.
    pub(crate) fn validate(&self) -> Result<(), XlsxError> {
        let range = self.formula();

        if self.first_row > self.last_row {
            return Err(XlsxError::ChartError(format!(
                "Chart series range '{range}' has a first row greater than the last row"
            )));
        }

        if self.first_col > self.last_col {
            return Err(XlsxError::ChartError(format!(
                "Chart series range '{range}' has a first column greater than the last column"
            )));
        }

        if self.first_row >= ROW_MAX || self.last_row >= ROW_MAX {
            return Err(XlsxError::ChartError(format!(
                "Chart series range '{range}' has a first row greater than Excel limit of 1048576"
            )));
        }

        if self.first_col >= COL_MAX || self.last_col >= COL_MAX {
            return Err(XlsxError::ChartError(
                format!("Chart series range '{range}' has a first column greater than Excel limit of XFD/16384"),
            ));
        }

        Ok(())
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
/// The `ChartType` enum define the type of a Chart object.
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_area.png">
    Area,

    /// A stacked Area chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_area_stacked.png">
    AreaStacked,

    /// A percent stacked Area chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_area_percent_stacked.png">
    AreaPercentStacked,

    /// A Bar (horizontal histogram) chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_bar.png">
    Bar,

    /// A stacked Bar chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_bar_stacked.png">
    BarStacked,

    /// A percent stacked Bar chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_bar_percent_stacked.png">
    BarPercentStacked,

    /// A Column (vertical histogram) chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_column.png">
    Column,

    /// A stacked Column chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_column_stacked.png">
    ColumnStacked,

    /// A percent stacked Column chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_column_percent_stacked.png">
    ColumnPercentStacked,

    /// A Doughnut chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_doughnut.png">
    Doughnut,

    /// An Line chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_line.png">
    Line,

    /// A stacked Line chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_line_stacked.png">
    LineStacked,

    /// A percent stacked Line chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_line_percent_stacked.png">
    LinePercentStacked,

    /// A Pie chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_pie.png">
    Pie,

    /// A Radar chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_radar.png">
    Radar,

    /// A Radar with markers chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_radar_with_markers.png">
    RadarWithMarkers,

    /// A filled Radar chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_radar_filled.png">
    RadarFilled,

    /// A Scatter chart type. Scatter charts, and their variant, are the only
    /// type that have values (as opposed to categories) for their X-Axis. The
    /// default scatter chart in Excel has markers for each point but no
    /// connecting lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_scatter.png">
    Scatter,

    /// A Scatter chart type where the points are connected by straight lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_scatter_straight.png">
    ScatterStraight,

    /// A Scatter chart type where the points have markers and are connected by
    /// straight lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_scatter_straight_with_markers.png">
    ScatterStraightWithMarkers,

    /// A Scatter chart type where the points are connected by smoothed lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_scatter_smooth.png">
    ScatterSmooth,

    /// A Scatter chart type where the points have markers and are connected by
    /// smoothed lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_scatter_smooth_with_markers.png">
    ScatterSmoothWithMarkers,
}

/// A struct to represent a Chart title.
#[derive(Clone)]
pub struct ChartTitle {
    pub(crate) range: ChartRange,
    pub(crate) cache_data: ChartSeriesCacheData,
    pub(crate) format: ChartFormat,
    name: String,
    hidden: bool,
    is_horizontal: bool,
}

impl ChartTitle {
    pub(crate) fn new() -> ChartTitle {
        ChartTitle {
            range: ChartRange::new_from_range("", 0, 0, 0, 0),
            cache_data: ChartSeriesCacheData::new(),
            format: ChartFormat::new(),
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
    /// #     // Create a new chart.
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
    /// series and a user defined series name. The `set_hidden()` method turns
    /// this default title off.
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
    /// #     // Create a new chart.
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
    ///  The output file is shown below. Note how there is a default title of
    /// "Yearly results" in the `"=SERIES("Yearly results",,Sheet1!$A$1:$A$3,1)"`
    /// formula in Excel but that it isn't displayed on the chart.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_title_set_hidden.png">
    ///
    pub fn set_hidden(&mut self) -> &mut ChartTitle {
        self.hidden = true;
        self
    }

    /// Set the formatting properties for a chart title.
    ///
    /// Set the formatting properties for a chart title via a [`ChartFormat`]
    /// object or a sub struct that implements [`IntoChartFormat`].
    ///
    /// The formatting that can be applied via a [`ChartFormat`] object are:
    ///
    /// - `no_fill`: Turn of the fill for the chart object.
    /// - `solid_fill`: Set the [`ChartSolidFill`] properties.
    /// - `pattern_fill`: Set the [`ChartPatternFill`] properties.
    /// - `no_line`: Turn off the line/border for the chart object.
    /// - `line`: Set the [`ChartLine`] properties.
    ///
    /// # Arguments
    ///
    /// `format`: A [`ChartFormat`] struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// [`IntoChartFormat`] for details.
    ///
    pub fn set_format<T>(&mut self, format: T) -> &mut ChartTitle
    where
        T: IntoChartFormat,
    {
        self.format = format.new_chart_format();
        self
    }
}

/// A struct to represent a Chart marker.
///
/// The [`ChartMarker`] struct represents the properties of a marker on a Line,
/// Scatter or Radar chart. In Excel a marker is a shape that represents a data
/// point in a chart series.
///
/// # Examples
///
/// An example of adding markers to a line chart.
///
/// ```
/// # // This code is available in examples/doc_chart_marker.rs
/// #
/// # use rust_xlsxwriter::{
/// #     Chart, ChartFormat, ChartMarker, ChartMarkerType, ChartSolidFill, ChartType, Workbook,
/// #     XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the chart.
/// #     worksheet.write(0, 0, 10)?;
/// #     worksheet.write(1, 0, 40)?;
/// #     worksheet.write(2, 0, 50)?;
/// #     worksheet.write(3, 0, 20)?;
/// #     worksheet.write(4, 0, 10)?;
/// #     worksheet.write(5, 0, 50)?;
/// #
/// #     // Create a new chart.
///     let mut chart = Chart::new(ChartType::Line);
///
///     // Add a data series with formatting.
///     chart
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_marker(
///             ChartMarker::new()
///                 .set_type(ChartMarkerType::Square)
///                 .set_size(10)
///                 .set_format(
///                     ChartFormat::new().set_solid_fill(
///                         ChartSolidFill::new().set_color("#FF0000")),
///                 ),
///         );
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
/// <img src="https://rustxlsxwriter.github.io/images/chart_marker.png">
///
#[derive(Clone)]
pub struct ChartMarker {
    pub(crate) automatic: bool,
    pub(crate) none: bool,
    pub(crate) size: u8,
    pub(crate) marker_type: Option<ChartMarkerType>,
    pub(crate) format: ChartFormat,
}

#[allow(clippy::new_without_default)]
impl ChartMarker {
    /// Create a new `ChartMarker` object to represent a Chart marker.
    ///
    pub fn new() -> ChartMarker {
        ChartMarker {
            automatic: false,
            none: false,
            marker_type: None,
            size: 0,
            format: ChartFormat::new(),
        }
    }

    /// Set the automatic/default marker type.
    ///
    /// Allow the marker type to be set automatically by Excel.
    ///
    /// # Examples
    ///
    /// An example of adding automatic markers to a line chart.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_marker_set_automatic.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartMarker, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_marker(ChartMarker::new().set_automatic());
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_set_automatic.png">
    ///
    pub fn set_automatic(&mut self) -> &mut ChartMarker {
        self.automatic = true;
        self
    }

    /// Turn off/hide a chart marker.
    ///
    /// This method can be use to turn off markers for an individual data series
    /// in a chart that has default markers for all series.
    ///
    pub fn set_none(&mut self) -> &mut ChartMarker {
        self.none = true;
        self
    }

    /// Set the type of the marker.
    ///
    /// Change the default type of the marker to one of the shapes supported by
    /// Excel.
    ///
    /// # Arguments
    ///
    /// `marker_type`: a [`ChartMarkerType`] enum value.
    ///
    /// # Examples
    ///
    /// An example of adding markers to a line chart with user defined marker
    /// types.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_marker_set_type.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartMarker, ChartMarkerType, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_marker(ChartMarker::new().set_type(ChartMarkerType::Circle));
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
    /// src="https://rustxlsxwriter.github.io/images/chart_marker_set_type.png">
    ///
    pub fn set_type(&mut self, marker_type: ChartMarkerType) -> &mut ChartMarker {
        self.marker_type = Some(marker_type);
        self.automatic = false;
        self
    }

    /// Set the size of the marker.
    ///
    /// Change the default size of the marker.
    ///
    ///
    ///
    pub fn set_size(&mut self, size: u8) -> &mut ChartMarker {
        if (2..=72).contains(&size) {
            self.size = size;
            self.automatic = false;
        }
        self
    }

    /// Set the formatting properties for a chart marker.
    ///
    /// Set the formatting properties for a chart marker via a [`ChartFormat`]
    /// object or a sub struct that implements [`IntoChartFormat`].
    ///
    /// The formatting that can be applied via a [`ChartFormat`] object are:
    ///
    /// - `no_fill`: Turn of the fill for the chart object.
    /// - `solid_fill`: Set the [`ChartSolidFill`] properties.
    /// - `pattern_fill`: Set the [`ChartPatternFill`] properties.
    /// - `no_line`: Turn off the line/border for the chart object.
    /// - `line`: Set the [`ChartLine`] properties.
    ///
    /// # Arguments
    ///
    /// `format`: A [`ChartFormat`] struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// [`IntoChartFormat`] for details.
    ///
    pub fn set_format<T>(&mut self, format: T) -> &mut ChartMarker
    where
        T: IntoChartFormat,
    {
        self.format = format.new_chart_format();
        self
    }
}

/// Enum to define the Chart marker types.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ChartMarkerType {
    /// Square marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_square.png">
    Square,

    /// Diamond marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_diamond.png">
    Diamond,

    /// Triangle marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_triangle.png">
    Triangle,

    /// X shape marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_x.png">
    X,

    /// Star (X overlaid on vertical dash) marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_star.png">
    Star,

    /// Short dash marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_short_dash.png">
    ShortDash,

    /// Long dash marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_long_dash.png">
    LongDash,

    /// Circle marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_circle.png">
    Circle,

    /// Plus sign marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_plus_sign.png">
    PlusSign,
}

impl ToString for ChartMarkerType {
    fn to_string(&self) -> String {
        match self {
            ChartMarkerType::X => "x".to_string(),
            ChartMarkerType::Star => "star".to_string(),
            ChartMarkerType::Circle => "circle".to_string(),
            ChartMarkerType::Square => "square".to_string(),
            ChartMarkerType::Diamond => "diamond".to_string(),
            ChartMarkerType::PlusSign => "plus".to_string(),
            ChartMarkerType::Triangle => "triangle".to_string(),
            ChartMarkerType::LongDash => "long_dash".to_string(),
            ChartMarkerType::ShortDash => "short_dash".to_string(),
        }
    }
}

/// A struct to represent a Chart point.
///
/// The [`ChartPoint`] struct represents a "point" in a data series which is the
/// element you get in Excel if you right click on an individual data point or
/// segment and select "Format Data Point".
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_point_dialog.png">
///
/// The meaning of "point" varies between chart types. For a Line chart a point
/// is a line segment; in a Column chart a point is a an individual bar; and in
/// a Pie chart a point is a pie segment.
///
/// Chart points are most commonly used for Pie and Doughnut charts to format
/// individual segments of the chart. In all other chart types the formatting
/// happens at the chart series level.
///
/// # Examples
///
/// An example of formatting the individual segments of a Pie chart.
///
/// ```
/// # // This code is available in examples/doc_chart_set_points.rs
/// #
/// # use rust_xlsxwriter::{
/// #     Chart, ChartFormat, ChartPoint, ChartSolidFill, ChartType, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the chart.
/// #     worksheet.write(0, 0, 15)?;
/// #     worksheet.write(1, 0, 15)?;
/// #     worksheet.write(2, 0, 30)?;
/// #
/// #     // Some point object with formatting to use in the Pie chart.
/// #     let points = vec![
/// #         ChartPoint::new().set_format(
/// #             ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FF0000")),
/// #         ),
/// #         ChartPoint::new().set_format(
/// #             ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FFC000")),
/// #         ),
/// #         ChartPoint::new().set_format(
/// #             ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FFFF00")),
/// #         ),
/// #     ];
/// #
/// #     // Create a simple Pie chart.
///     let mut chart = Chart::new(ChartType::Pie);
///
///     // Add a data series with formatting.
///     chart
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$3")
///         .set_points(&points);
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
/// <img src="https://rustxlsxwriter.github.io/images/chart_set_points.png">
///
#[derive(Clone)]
pub struct ChartPoint {
    pub(crate) format: ChartFormat,
}

impl Default for ChartPoint {
    fn default() -> Self {
        Self::new()
    }
}

impl ChartPoint {
    /// Create a new `ChartPoint` object to represent a Chart point.
    ///
    pub fn new() -> ChartPoint {
        ChartPoint {
            format: ChartFormat::new(),
        }
    }

    /// Set the formatting properties for a chart point.
    ///
    /// Set the formatting properties for a chart point via a [`ChartFormat`]
    /// object or a sub struct that implements [`IntoChartFormat`].
    ///
    /// The formatting that can be applied via a [`ChartFormat`] object are:
    ///
    /// - `no_fill`: Turn of the fill for the chart object.
    /// - `solid_fill`: Set the [`ChartSolidFill`] properties.
    /// - `pattern_fill`: Set the [`ChartPatternFill`] properties.
    /// - `no_line`: Turn off the line/border for the chart object.
    /// - `line`: Set the [`ChartLine`] properties.
    ///
    /// # Arguments
    ///
    /// `format`: A [`ChartFormat`] struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// [`IntoChartFormat`] for details.
    ///
    pub fn set_format<T>(mut self, format: T) -> ChartPoint
    where
        T: IntoChartFormat,
    {
        self.format = format.new_chart_format();
        self
    }

    pub(crate) fn is_not_default(&self) -> bool {
        self.format.has_formatting()
    }
}

/// A struct to represent a Chart axis.
#[derive(Clone)]
pub struct ChartAxis {
    axis_type: ChartAxisType,
    axis_position: ChartAxisPosition,
    pub(crate) title: ChartTitle,
    pub(crate) format: ChartFormat,
}

impl ChartAxis {
    pub(crate) fn new() -> ChartAxis {
        ChartAxis {
            axis_type: ChartAxisType::Value,
            axis_position: ChartAxisPosition::Bottom,
            title: ChartTitle::new(),
            format: ChartFormat::new(),
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
    /// #     // Create a new chart.
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

    /// Set the formatting properties for a chart axis.
    ///
    /// Set the formatting properties for a chart axis via a [`ChartFormat`]
    /// object or a sub struct that implements [`IntoChartFormat`].
    ///
    /// The formatting that can be applied via a [`ChartFormat`] object are:
    ///
    /// - `no_fill`: Turn of the fill for the chart object.
    /// - `solid_fill`: Set the [`ChartSolidFill`] properties.
    /// - `pattern_fill`: Set the [`ChartPatternFill`] properties.
    /// - `no_line`: Turn off the line/border for the chart object.
    /// - `line`: Set the [`ChartLine`] properties.
    ///
    /// # Arguments
    ///
    /// `format`: A [`ChartFormat`] struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// [`IntoChartFormat`] for details.
    ///
    pub fn set_format<T>(&mut self, format: T) -> &mut ChartAxis
    where
        T: IntoChartFormat,
    {
        self.format = format.new_chart_format();
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

/// A struct to represent a Chart legend.
///
/// The `ChartLegend` struct is a representation of a legend on an Excel chart.
/// The legend is a rectangular box that identifies the name and color of each
/// of the series in the chart.
///
/// `ChartLegend` can be used to configure properties of the chart legend and is
/// usually obtained via the [`chart.legend()`][Chart::legend] method.
///
/// # Examples
///
/// An example of getting the chart legend object and setting some of its
/// properties.
///
/// ```
/// # // This code is available in examples/doc_chart_legend.rs
/// #
/// # use rust_xlsxwriter::{Chart, ChartLegendPosition, ChartType, Workbook, XlsxError};
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
/// #     worksheet.write(1, 1, 35)?;
/// #     worksheet.write(2, 1, 45)?;
/// #
/// #     // Create a new chart.
///     let mut chart = Chart::new(ChartType::Column);
///
///     // Add a data series using Excel formula syntax to describe the range.
///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
///     chart.add_series().set_values("Sheet1!$B$1:$B$3");
///
///     // Turn on the chart legend and place it at the bottom of the chart.
///     chart.legend().set_position(ChartLegendPosition::Bottom);
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
/// <img src="https://rustxlsxwriter.github.io/images/chart_legend.png">
///
#[derive(Clone)]
pub struct ChartLegend {
    position: ChartLegendPosition,
    hidden: bool,
    has_overlay: bool,
    pub(crate) format: ChartFormat,
}

impl ChartLegend {
    pub(crate) fn new() -> ChartLegend {
        ChartLegend {
            position: ChartLegendPosition::Right,
            hidden: false,
            has_overlay: false,
            format: ChartFormat::new(),
        }
    }

    /// Hide the legend for a Chart.
    ///
    /// The legend if usually on by default for an Excel chart. The
    /// `set_hidden()` method can be used to turn it off.
    ///
    /// # Examples
    ///
    /// An example of hiding a default chart legend.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_legend_set_hidden.rs
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
    /// #     // Create a new chart.
    /// #     let mut chart = Chart::new(ChartType::Column);
    /// #
    /// #     // Add a data series using Excel formula syntax to describe the range.
    /// #     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    /// #
    /// #     // Hide the chart legend.
    /// #     chart.legend().set_hidden();
    /// #
    /// #     // Add the chart to the worksheet.
    /// #     worksheet.insert_chart(0, 2, &chart)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_legend_set_hidden.png">
    ///
    pub fn set_hidden(&mut self) -> &mut ChartLegend {
        self.hidden = true;
        self
    }

    /// Set the chart legend position.
    ///
    /// Set the position of the legend on the chart. The available positions in
    /// Excel are:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/legend_position.png">
    ///
    /// The equivalent positions in rust_xlsxwriter charts are defined by
    /// [`ChartLegendPosition`]. The default chart position in Excel is to have
    /// the legend at the right.
    ///
    /// # Examples
    ///
    /// An example of setting the position of the chart legend.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_legend.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartLegendPosition, ChartType, Workbook, XlsxError};
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
    /// #     worksheet.write(1, 1, 35)?;
    /// #     worksheet.write(2, 1, 45)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///     chart.add_series().set_values("Sheet1!$B$1:$B$3");
    ///
    ///     // Turn on the chart legend and place it at the bottom of the chart.
    ///     chart.legend().set_position(ChartLegendPosition::Bottom);
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_legend.png">
    ///
    pub fn set_position(&mut self, position: ChartLegendPosition) -> &mut ChartLegend {
        self.position = position;
        self
    }

    /// Set the chart legend as overlaid on the chart.
    ///
    /// In the Excel "Format Legend" dialog there is a default option of "Show
    /// the legend without overlapping the chart":
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/legend_position.png">
    ///
    /// This can be turned off using the `set_overlay()` method.
    ///
    /// # Examples
    ///
    /// An example of overlaying the chart legend on the plot area.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_legend_set_overlay.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, ChartLegendPosition, Workbook, XlsxError};
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
    /// #     // Create a new chart.
    /// #     let mut chart = Chart::new(ChartType::Column);
    /// #
    /// #     // Add a data series using Excel formula syntax to describe the range.
    /// #     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    /// #
    /// #     // Turn on the chart legend and place it at the top of the chart.
    /// #     chart.legend().set_position(ChartLegendPosition::Top);
    /// #
    /// #     // Overlay the chart legend on the plot area.
    /// #     chart.legend().set_overlay();
    /// #
    /// #     // Add the chart to the worksheet.
    /// #     worksheet.insert_chart(0, 2, &chart)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_legend_set_overlay.png">
    ///
    pub fn set_overlay(&mut self) -> &mut ChartLegend {
        self.has_overlay = true;
        self
    }

    /// Set the formatting properties for a chart legend.
    ///
    /// Set the formatting properties for a chart legend via a [`ChartFormat`]
    /// object or a sub struct that implements [`IntoChartFormat`].
    ///
    /// The formatting that can be applied via a [`ChartFormat`] object are:
    ///
    /// - `no_fill`: Turn of the fill for the chart object.
    /// - `solid_fill`: Set the [`ChartSolidFill`] properties.
    /// - `pattern_fill`: Set the [`ChartPatternFill`] properties.
    /// - `no_line`: Turn off the line/border for the chart object.
    /// - `line`: Set the [`ChartLine`] properties.
    ///
    /// # Arguments
    ///
    /// `format`: A [`ChartFormat`] struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// [`IntoChartFormat`] for details.
    ///
    pub fn set_format<T>(&mut self, format: T) -> &mut ChartLegend
    where
        T: IntoChartFormat,
    {
        self.format = format.new_chart_format();
        self
    }
}

/// Enum used to specify the position of the Chart legend.
///
/// Excel allows the following positions of the chart legend:
///
/// <img src="https://rustxlsxwriter.github.io/images/legend_position.png">
///
/// These positions can be set using the
/// [`chart.legend().set_position()`](ChartLegend::set_position) method and the
/// `ChartLegendPosition` enum values.
///
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ChartLegendPosition {
    /// Chart legend positioned at the right side. The default.
    Right,

    /// Chart legend positioned at the left side.
    Left,

    /// Chart legend positioned at the top.
    Top,

    /// Chart legend positioned at the bottom.
    Bottom,

    /// Chart legend positioned at the top right.
    TopRight,
}

impl ToString for ChartLegendPosition {
    fn to_string(&self) -> String {
        match self {
            ChartLegendPosition::Top => "t".to_string(),
            ChartLegendPosition::Left => "l".to_string(),
            ChartLegendPosition::Right => "r".to_string(),
            ChartLegendPosition::Bottom => "b".to_string(),
            ChartLegendPosition::TopRight => "tr".to_string(),
        }
    }
}

#[derive(Clone)]
/// A struct to represent formatting for various Chart objects.
///
/// Excel uses a standard formatting dialog for the elements of a chart such as
/// data series, the plot area, the chart area, the legend or individual points.
/// It looks like this:
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_format_dialog.png">
///
/// The [`ChartFormat`] struct represents many of these format options and just
/// like Excel it offers a similar formatting interface for a number of the
/// chart sub-elements supported by rust_xlsxwriter.
///
/// The [`ChartFormat`] struct is accessed by using the `set_format()` method of a
/// chart element to obtain a reference to the formatting struct for that
/// element. After that it can be used to apply formatting such as:
///
/// - `no_fill`: Turn of the fill for the chart object.
/// - `solid_fill`: Set the [`ChartSolidFill`] properties.
/// - `pattern_fill`: Set the [`ChartPatternFill`] properties.
/// - `no_line`: Turn off the line/border for the chart object.
/// - `line`: Set the [`ChartLine`] properties for lines or borders.
///
/// # Examples
///
/// A example of accessing the [`ChartFormat`] for data series in a chart and
/// using them to apply formatting.
///
/// ```
/// # // This code is available in examples/app_chart_pattern.rs
/// #
/// # use rust_xlsxwriter::*;
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #     let bold = Format::new().set_bold();
/// #
/// #     // Add the worksheet data that the charts will refer to.
/// #     worksheet.write_with_format(0, 0, "Shingle", &bold)?;
/// #     worksheet.write_with_format(0, 1, "Brick", &bold)?;
/// #
/// #     let data = [[105, 150, 130, 90], [50, 120, 100, 110]];
/// #     for (col_num, col_data) in data.iter().enumerate() {
/// #         for (row_num, row_data) in col_data.iter().enumerate() {
/// #             worksheet.write(row_num as u32 + 1, col_num as u16, *row_data)?;
/// #         }
/// #     }
/// #
/// #     // Create a new column chart.
///     let mut chart = Chart::new(ChartType::Column);
///
///     // Configure the first data series and add fill patterns.
///     chart
///         .add_series()
///         .set_name("Sheet1!$A$1")
///         .set_values("Sheet1!$A$2:$A$5")
///         .set_gap(70)
///         .set_format(
///             ChartFormat::new()
///                 .set_pattern_fill(
///                     ChartPatternFill::new()
///                         .set_pattern(ChartPatternFillType::Shingle)
///                         .set_foreground_color("#804000")
///                         .set_background_color("#C68C53"),
///                 )
///                 .set_border(ChartLine::new().set_color("#804000")),
///         );
///
///     chart
///         .add_series()
///         .set_name("Sheet1!$B$1")
///         .set_values("Sheet1!$B$2:$B$5")
///         .set_format(
///             ChartFormat::new()
///                 .set_pattern_fill(
///                     ChartPatternFill::new()
///                         .set_pattern(ChartPatternFillType::HorizontalBrick)
///                         .set_foreground_color("#B30000")
///                         .set_background_color("#FF6666"),
///                 )
///                 .set_border(ChartLine::new().set_color("#B30000")),
///         );
///
///     // Add a chart title and some axis labels.
///     chart.title().set_name("Cladding types");
///     chart.x_axis().set_name("Region");
///     chart.y_axis().set_name("Number of houses");
///
///     // Add the chart to the worksheet.
///     worksheet.insert_chart(1, 3, &chart)?;
///
///     workbook.save("chart_pattern.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/app_chart_pattern.png">
///
pub struct ChartFormat {
    no_fill: bool,
    no_line: bool,
    line: Option<ChartLine>,
    solid_fill: Option<ChartSolidFill>,
    pattern_fill: Option<ChartPatternFill>,
}

impl ChartFormat {
    /// Create a new `ChartFormat` instance to set formatting for a chart element.
    ///
    #[allow(clippy::new_without_default)]
    pub fn new() -> ChartFormat {
        ChartFormat {
            no_fill: false,
            no_line: false,
            line: None,
            solid_fill: None,
            pattern_fill: None,
        }
    }

    /// Set the line formatting for a chart element.
    ///
    /// See the [`ChartLine`] struct for details on the line properties that can
    /// be set.
    ///
    pub fn set_line(&mut self, line: &ChartLine) -> &mut ChartFormat {
        self.line = Some(line.clone());
        self
    }

    /// Set the border formatting for a chart element.
    ///
    /// See the [`ChartLine`] struct for details on the border properties that
    /// can be set.
    ///
    /// # Examples
    ///
    /// An example of formatting a line/border in a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_line_formatting.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartFormat, ChartLine, ChartLineDashType, ChartType, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ChartFormat::new().set_line(
    ///                 ChartLine::new()
    ///                     .set_color("#FF9900")
    ///                     .set_width(5.25)
    ///                     .set_dash_type(ChartLineDashType::SquareDot)
    ///                     .set_transparency(70),
    ///             ),
    ///         );
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_formatting.png">
    ///
    pub fn set_border(&mut self, line: &ChartLine) -> &mut ChartFormat {
        self.set_line(line)
    }

    /// Turn off the line property for a chart element.
    ///
    /// The line property for a chart element can be turned off if you wish to
    /// hide it.
    ///
    /// # Examples
    ///
    /// An example of turning off a default line in a chart format.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_format_set_no_line.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFormat, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 1)?;
    /// #     worksheet.write(1, 0, 2)?;
    /// #     worksheet.write(2, 0, 3)?;
    /// #     worksheet.write(3, 0, 4)?;
    /// #     worksheet.write(4, 0, 5)?;
    /// #     worksheet.write(5, 0, 6)?;
    /// #     worksheet.write(0, 1, 10)?;
    /// #     worksheet.write(1, 1, 40)?;
    /// #     worksheet.write(2, 1, 50)?;
    /// #     worksheet.write(3, 1, 20)?;
    /// #     worksheet.write(4, 1, 10)?;
    /// #     worksheet.write(5, 1, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::ScatterStraightWithMarkers);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_categories("Sheet1!$A$1:$A$6")
    ///         .set_values("Sheet1!$B$1:$B$6")
    ///         .set_format(ChartFormat::new().set_no_line());
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_format_set_no_line.png">
    ///
    pub fn set_no_line(&mut self) -> &mut ChartFormat {
        self.no_line = true;
        self
    }

    /// Turn off the border property for a chart element.
    ///
    /// The border property for a chart element can be turned off if you wish to
    /// hide it.
    ///
    /// # Examples
    ///
    /// An example of turning off the border of a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_format_set_no_border.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFormat, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(ChartFormat::new().set_no_border());
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_format_set_no_border.png">
    ///
    pub fn set_no_border(&mut self) -> &mut ChartFormat {
        self.set_no_line()
    }

    /// Turn off the fill property for a chart element.
    ///
    /// The fill property for a chart element can be turned off if you wish to
    /// hide it and display only the border (if set).
    ///
    /// # Examples
    ///
    /// An example of turning off the fill of a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_format_set_no_fill.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFormat, ChartLine, ChartType, Workbook, XlsxColor, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ChartFormat::new()
    ///                 .set_border(ChartLine::new().set_color(XlsxColor::Black))
    ///                 .set_no_fill(),
    ///         );
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_format_set_no_fill.png">
    ///
    pub fn set_no_fill(&mut self) -> &mut ChartFormat {
        self.no_fill = true;
        self
    }

    /// Set the solid fill formatting for a chart element.
    ///
    /// See the [`ChartSolidFill`] struct for details on the solid fill
    /// properties that can be set.
    ///
    /// # Examples
    ///
    /// An example of setting a solid fill for a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_solid_fill.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFormat, ChartSolidFill, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ChartFormat::new().set_solid_fill(
    ///                 ChartSolidFill::new()
    ///                     .set_color("#FF9900")
    ///                     .set_transparency(60),
    ///             ),
    ///         );
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_solid_fill.png">
    ///
    pub fn set_solid_fill(&mut self, fill: &ChartSolidFill) -> &mut ChartFormat {
        self.solid_fill = Some(fill.clone());
        self
    }

    /// Set the pattern fill formatting for a chart element.
    ///
    /// See the [`ChartPatternFill`] struct for details on the pattern fill
    /// properties that can be set.
    ///
    /// # Examples
    ///
    /// An example of setting a pattern fill for a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_pattern_fill.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartFormat, ChartPatternFill, ChartPatternFillType, ChartType, Workbook, XlsxColor,
    /// #     XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ChartFormat::new().set_pattern_fill(
    ///                 ChartPatternFill::new()
    ///                     .set_pattern(ChartPatternFillType::Dotted20Percent)
    ///                     .set_background_color(XlsxColor::Yellow)
    ///                     .set_foreground_color(XlsxColor::Red),
    ///             ),
    ///         );
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill.png">
    ///
    pub fn set_pattern_fill(&mut self, fill: &ChartPatternFill) -> &mut ChartFormat {
        self.pattern_fill = Some(fill.clone());
        self
    }

    // Check if formatting has been set for the struct.
    fn has_formatting(&self) -> bool {
        self.line.is_some()
            || self.solid_fill.is_some()
            || self.pattern_fill.is_some()
            || self.no_fill
            || self.no_line
    }
}

/// A struct to represent a Chart line/border.
///
/// The [`ChartLine`] struct represents the formatting properties for a line or
/// border for a Chart element. It is a sub property of the [`ChartFormat`]
/// struct and is used with the [`ChartFormat::set_line()`](ChartFormat::set_line)
/// or [`ChartFormat::set_border()`](ChartFormat::set_border) methods.
///
/// Excel uses the element names "Line" and "Border" depending on the context.
/// For a Line chart the line is represented by a line property but for a Column
/// chart the line becomes the border. Both of these share the same properties
/// and are both represented in rust_xlsxwriter by the [`ChartLine`] struct.
///
/// # Examples
///
/// An example of formatting a line/border in a chart element.
///
/// ```
/// # // This code is available in examples/doc_chart_line_formatting.rs
/// #
/// # use rust_xlsxwriter::{
/// #     Chart, ChartFormat, ChartLine, ChartLineDashType, ChartType, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the chart.
/// #     worksheet.write(0, 0, 10)?;
/// #     worksheet.write(1, 0, 40)?;
/// #     worksheet.write(2, 0, 50)?;
/// #     worksheet.write(3, 0, 20)?;
/// #     worksheet.write(4, 0, 10)?;
/// #     worksheet.write(5, 0, 50)?;
/// #
/// #     // Create a new chart.
///     let mut chart = Chart::new(ChartType::Line);
///
///     // Add a data series with formatting.
///     chart
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_format(
///             ChartFormat::new().set_line(
///                 ChartLine::new()
///                     .set_color("#FF9900")
///                     .set_width(5.25)
///                     .set_dash_type(ChartLineDashType::SquareDot)
///                     .set_transparency(70),
///             ),
///         );
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
/// <img src="https://rustxlsxwriter.github.io/images/chart_line_formatting.png">
///
#[derive(Clone)]
pub struct ChartLine {
    color: XlsxColor,
    width: Option<f64>,
    transparency: u8,
    dash_type: ChartLineDashType,
    hidden: bool,
}

impl ChartLine {
    /// Create a new `ChartLine` object to represent a Chart line/border.
    ///
    #[allow(clippy::new_without_default)]
    pub fn new() -> ChartLine {
        ChartLine {
            color: XlsxColor::Default,
            width: None,
            transparency: 0,
            dash_type: ChartLineDashType::Solid,
            hidden: false,
        }
    }

    /// Set the color of a line/border.
    ///
    /// # Arguments
    ///
    /// * `color` - The color property defined by a [`XlsxColor`] enum value or
    ///   a type that implements the [`IntoColor`] trait.
    ///
    /// # Examples
    ///
    /// An example of formatting the line color in a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_line_set_color.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFormat, ChartLine, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(ChartFormat::new().set_line(ChartLine::new().set_color("#FF9900")));
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_set_color.png">
    ///
    pub fn set_color<T>(&mut self, color: T) -> &mut ChartLine
    where
        T: IntoColor,
    {
        let color = color.new_color();
        if color.is_valid() {
            self.color = color;
        }

        self
    }

    /// Set the width of the line or border.
    ///
    /// # Arguments
    ///
    /// * `width` - The width should be specified in increments of 0.25 of a
    /// point as in Excel. The width can be an number type that convert [`Into`]
    /// [`f64`].
    ///
    /// # Examples
    ///
    /// An example of formatting the line width in a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_line_set_width.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFormat, ChartLine, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(ChartFormat::new().set_line(ChartLine::new().set_width(10.0)));
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
    /// src="https://rustxlsxwriter.github.io/images/chart_line_set_width.png">
    ///
    pub fn set_width<T>(&mut self, width: T) -> &mut ChartLine
    where
        T: Into<f64>,
    {
        let width = width.into();
        if width <= 1584.0 {
            self.width = Some(width);
        }

        self
    }

    /// Set the dash type of the line or border.
    ///
    /// # Arguments
    ///
    /// * `dash_type` - A [`ChartLineDashType`] enum value.
    ///
    /// # Examples
    ///
    /// An example of formatting the line dash type in a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_line_set_dash_type.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartFormat, ChartLine, ChartLineDashType, ChartType, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ChartFormat::new().set_line(ChartLine::new().set_dash_type(ChartLineDashType::DashDot)),
    ///         );
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_set_dash_type.png">
    ///
    pub fn set_dash_type(&mut self, dash_type: ChartLineDashType) -> &mut ChartLine {
        self.dash_type = dash_type;
        self
    }

    /// Set the transparency of a line/border.
    ///
    /// Set the transparency of a line/border for a Chart element. You must also
    /// specify a line color in order for the transparency to be applied.
    ///
    /// # Arguments
    ///
    /// * `transparency` - The color transparency in the range 0 <= transparency
    ///   <= 100. The default value is 0.
    ///
    /// # Examples
    ///
    /// An example of formatting the line transparency in a chart element. Note, you
    /// must set also set a color in order to set the transparency.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_line_set_transparency.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFormat, ChartLine, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ChartFormat::new().set_line(ChartLine::new().set_color("#FF9900").set_transparency(50)),
    ///         );
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_set_transparency.png">
    ///
    pub fn set_transparency(&mut self, transparency: u8) -> &mut ChartLine {
        if transparency <= 100 {
            self.transparency = transparency;
        }

        self
    }

    // Internal method for some chart types such as Scatter that set a line
    // width but also set the line hidden.
    pub(crate) fn set_hidden(&mut self) -> &mut ChartLine {
        self.hidden = true;
        self
    }
}

/// A struct to represent a the solid fill for a Chart element.
///
/// The [`ChartSolidFill`] struct represents the formatting properties for the
/// solid fill of a Chart element. In Excel a solid fill is a single color fill
/// without a pattern or gradient.
///
/// `ChartSolidFill` is a sub property of the [`ChartFormat`] struct and is used
/// with the [`ChartFormat::set_solid_fill()`](ChartFormat::set_solid_fill) method.
///
/// # Examples
///
/// An example of setting a solid fill for a chart element.
///
/// ```
/// # // This code is available in examples/doc_chart_solid_fill.rs
/// #
/// # use rust_xlsxwriter::{Chart, ChartFormat, ChartSolidFill, ChartType, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the chart.
/// #     worksheet.write(0, 0, 10)?;
/// #     worksheet.write(1, 0, 40)?;
/// #     worksheet.write(2, 0, 50)?;
/// #     worksheet.write(3, 0, 20)?;
/// #     worksheet.write(4, 0, 10)?;
/// #     worksheet.write(5, 0, 50)?;
/// #
/// #     // Create a new chart.
///     let mut chart = Chart::new(ChartType::Column);
///
///     // Add a data series with formatting.
///     chart
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_format(
///             ChartFormat::new().set_solid_fill(
///                 ChartSolidFill::new()
///                     .set_color("#FF9900")
///                     .set_transparency(60),
///             ),
///         );
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
/// <img src="https://rustxlsxwriter.github.io/images/chart_solid_fill.png">
///
#[derive(Clone)]
pub struct ChartSolidFill {
    color: XlsxColor,
    transparency: u8,
}

impl ChartSolidFill {
    /// Create a new `ChartSolidFill` object to represent a Chart solid fill.
    ///
    #[allow(clippy::new_without_default)]
    pub fn new() -> ChartSolidFill {
        ChartSolidFill {
            color: XlsxColor::Default,
            transparency: 0,
        }
    }

    /// Set the color of a solid fill.
    ///
    /// # Arguments
    ///
    /// * `color` - The color property defined by a [`XlsxColor`] enum value or
    ///   a type that implements the [`IntoColor`] trait.
    ///
    /// # Examples
    ///
    /// An example of setting a solid fill color for a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_solid_fill_set_color.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFormat, ChartSolidFill, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#B5A401")));
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_solid_fill_set_color.png">
    ///
    pub fn set_color<T>(&mut self, color: T) -> &mut ChartSolidFill
    where
        T: IntoColor,
    {
        let color = color.new_color();
        if color.is_valid() {
            self.color = color;
        }

        self
    }

    /// Set the transparency of a solid fill.
    ///
    /// Set the transparency of a solid fill color for a Chart element. You must
    /// also specify a line color in order for the transparency to be applied.
    ///
    /// # Arguments
    ///
    /// * `transparency` - The color transparency in the range 0 <= transparency
    ///   <= 100. The default value is 0.
    ///
    /// # Examples
    ///
    /// An example of setting a solid fill for a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_solid_fill.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFormat, ChartSolidFill, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ChartFormat::new().set_solid_fill(
    ///                 ChartSolidFill::new()
    ///                     .set_color("#FF9900")
    ///                     .set_transparency(60),
    ///             ),
    ///         );
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_solid_fill.png">
    ///
    pub fn set_transparency(&mut self, transparency: u8) -> &mut ChartSolidFill {
        if transparency <= 100 {
            self.transparency = transparency;
        }

        self
    }
}

/// A struct to represent a the pattern fill for a Chart element.
///
/// The [`ChartPatternFill`] struct represents the formatting properties for the
/// pattern fill of a Chart element. In Excel a pattern fill is comprised of a
/// simple pixelated pattern and background and foreground colors
///
/// `ChartPatternFill` is a sub property of the [`ChartFormat`] struct and is
/// used with the [`ChartFormat::set_pattern_fill()`](ChartFormat::set_pattern_fill)
/// method.
///
///
/// # Examples
///
/// An example of setting a pattern fill for a chart element.
///
/// ```
/// # // This code is available in examples/doc_chart_pattern_fill.rs
/// #
/// # use rust_xlsxwriter::{
/// #     Chart, ChartFormat, ChartPatternFill, ChartPatternFillType, ChartType, Workbook, XlsxColor, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the chart.
/// #     worksheet.write(0, 0, 10)?;
/// #     worksheet.write(1, 0, 40)?;
/// #     worksheet.write(2, 0, 50)?;
/// #     worksheet.write(3, 0, 20)?;
/// #     worksheet.write(4, 0, 10)?;
/// #     worksheet.write(5, 0, 50)?;
/// #
/// #     // Create a new chart.
///     let mut chart = Chart::new(ChartType::Column);
///
///     // Add a data series with formatting.
///     chart
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_format(
///             ChartFormat::new().set_pattern_fill(
///                 ChartPatternFill::new()
///                     .set_pattern(ChartPatternFillType::Dotted20Percent)
///                     .set_background_color(XlsxColor::Yellow)
///                     .set_foreground_color(XlsxColor::Red),
///             ),
///         );
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
/// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill.png">
///
#[derive(Clone)]
pub struct ChartPatternFill {
    background_color: XlsxColor,
    foreground_color: XlsxColor,
    pattern: ChartPatternFillType,
}

impl ChartPatternFill {
    /// Create a new `ChartPatternFill` object to represent a Chart pattern fill.
    ///
    #[allow(clippy::new_without_default)]
    pub fn new() -> ChartPatternFill {
        ChartPatternFill {
            background_color: XlsxColor::Default,
            foreground_color: XlsxColor::Default,
            pattern: ChartPatternFillType::Dotted5Percent,
        }
    }

    /// Set the pattern of a Chart pattern fill element.
    ///
    /// See the example above.
    ///
    /// # Arguments
    ///
    /// * `pattern` - The pattern property defined by a [`ChartPatternFillType`] enum value.
    ///
    ///
    /// # Examples
    ///
    /// An example of setting a pattern fill for a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_pattern_fill_set_pattern.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartFormat, ChartPatternFill, ChartPatternFillType, ChartType, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(ChartFormat::new().set_pattern_fill(
    ///             ChartPatternFill::new().set_pattern(ChartPatternFillType::DiagonalBrick),
    ///         ));
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_set_pattern.png">
    ///
    pub fn set_pattern(&mut self, pattern: ChartPatternFillType) -> &mut ChartPatternFill {
        self.pattern = pattern;
        self
    }

    /// Set the background color of a Chart pattern fill element.
    ///
    /// See the example above.
    ///
    /// # Arguments
    ///
    /// * `color` - The color property defined by a [`XlsxColor`] enum value or
    ///   a type that implements the [`IntoColor`] trait.
    ///
    /// # Examples
    ///
    /// An example of setting a pattern fill for a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_pattern_fill.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartFormat, ChartPatternFill, ChartPatternFillType, ChartType, Workbook, XlsxColor,
    /// #     XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ChartFormat::new().set_pattern_fill(
    ///                 ChartPatternFill::new()
    ///                     .set_pattern(ChartPatternFillType::Dotted20Percent)
    ///                     .set_background_color(XlsxColor::Yellow)
    ///                     .set_foreground_color(XlsxColor::Red),
    ///             ),
    ///         );
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill.png">
    ///
    pub fn set_background_color<T>(&mut self, color: T) -> &mut ChartPatternFill
    where
        T: IntoColor,
    {
        let color = color.new_color();
        if color.is_valid() {
            self.background_color = color;
        }

        self
    }

    /// Set the foreground color of a Chart pattern fill element.
    ///
    /// See the example above.
    ///
    /// # Arguments
    ///
    /// * `color` - The color property defined by a [`XlsxColor`] enum value or
    ///   a type that implements the [`IntoColor`] trait.
    ///
    pub fn set_foreground_color<T>(&mut self, color: T) -> &mut ChartPatternFill
    where
        T: IntoColor,
    {
        let color = color.new_color();
        if color.is_valid() {
            self.foreground_color = color;
        }

        self
    }
}

/// Enum to define the Chart line dash type.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ChartLineDashType {
    /// Solid - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_solid.png">
    Solid,

    /// Round dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_round_dot.png">
    RoundDot,

    /// Square dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_square_dot.png">
    SquareDot,

    /// Dash - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_dash.png">
    Dash,

    /// Dash dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_dash_dot.png">
    DashDot,

    /// Long dash - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash.png">
    LongDash,

    /// Long dash dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash_dot.png">
    LongDashDot,

    /// Long dash dot dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash_dot_dot.png">
    LongDashDotDot,
}

impl ToString for ChartLineDashType {
    fn to_string(&self) -> String {
        match self {
            ChartLineDashType::Dash => "dash".to_string(),
            ChartLineDashType::Solid => "solid".to_string(),
            ChartLineDashType::DashDot => "dashDot".to_string(),
            ChartLineDashType::LongDash => "lgDash".to_string(),
            ChartLineDashType::RoundDot => "sysDot".to_string(),
            ChartLineDashType::SquareDot => "sysDash".to_string(),
            ChartLineDashType::LongDashDot => "lgDashDot".to_string(),
            ChartLineDashType::LongDashDotDot => "lgDashDotDot".to_string(),
        }
    }
}

/// Enum to define the Chart pattern fill type.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ChartPatternFillType {
    /// Dotted 5 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_5_percent.png">
    Dotted5Percent,

    /// Dotted 10 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_10_percent.png">
    Dotted10Percent,

    /// Dotted 20 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_20_percent.png">
    Dotted20Percent,

    /// Dotted 25 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_25_percent.png">
    Dotted25Percent,

    /// Dotted 30 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_30_percent.png">
    Dotted30Percent,

    /// Dotted 40 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_40_percent.png">
    Dotted40Percent,

    /// Dotted 50 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_50_percent.png">
    Dotted50Percent,

    /// Dotted 60 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_60_percent.png">
    Dotted60Percent,

    /// Dotted 70 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_70_percent.png">
    Dotted70Percent,

    /// Dotted 75 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_75_percent.png">
    Dotted75Percent,

    /// Dotted 80 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_80_percent.png">
    Dotted80Percent,

    /// Dotted 90 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_90_percent.png">
    Dotted90Percent,

    /// Diagonal stripes light downwards - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_light_downwards.png">
    DiagonalStripesLightDownwards,

    /// Diagonal stripes light upwards - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_light_upwards.png">
    DiagonalStripesLightUpwards,

    /// Diagonal stripes dark downwards - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_dark_downwards.png">
    DiagonalStripesDarkDownwards,

    /// Diagonal stripes dark upwards - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_dark_upwards.png">
    DiagonalStripesDarkUpwards,

    /// Diagonal stripes wide downwards - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_wide_downwards.png">
    DiagonalStripesWideDownwards,

    /// Diagonal stripes wide upwards - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_wide_upwards.png">
    DiagonalStripesWideUpwards,

    /// Vertical stripes light - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_light.png">
    VerticalStripesLight,

    /// Horizontal stripes light - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_light.png">
    HorizontalStripesLight,

    /// Vertical stripes narrow - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_narrow.png">
    VerticalStripesNarrow,

    /// Horizontal stripes narrow - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_narrow.png">
    HorizontalStripesNarrow,

    /// Vertical stripes dark - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_dark.png">
    VerticalStripesDark,

    /// Horizontal stripes dark - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_dark.png">
    HorizontalStripesDark,

    /// Stripes backslashes - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_stripes_backslashes.png">
    StripesBackslashes,

    /// Stripes forward slashes - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_stripes_forward_slashes.png">
    StripesForwardSlashes,

    /// Horizontal stripes alternating - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_alternating.png">
    HorizontalStripesAlternating,

    /// Vertical stripes alternating - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_alternating.png">
    VerticalStripesAlternating,

    /// Small confetti - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_small_confetti.png">
    SmallConfetti,

    /// Large confetti - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_large_confetti.png">
    LargeConfetti,

    /// Zigzag - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_zigzag.png">
    Zigzag,

    /// Wave - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_wave.png">
    Wave,

    /// Diagonal brick - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_brick.png">
    DiagonalBrick,

    /// Horizontal brick - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_brick.png">
    HorizontalBrick,

    /// Weave - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_weave.png">
    Weave,

    /// Plaid - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_plaid.png">
    Plaid,

    /// Divot - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_divot.png">
    Divot,

    /// Dotted grid - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_grid.png">
    DottedGrid,

    /// Dotted diamond - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_diamond.png">
    DottedDiamond,

    /// Shingle - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_shingle.png">
    Shingle,

    /// Trellis - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_trellis.png">
    Trellis,

    /// Sphere - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_sphere.png">
    Sphere,

    /// Small grid - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_small_grid.png">
    SmallGrid,

    /// Large grid - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_large_grid.png">
    LargeGrid,

    /// Small checkerboard - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_small_checkerboard.png">
    SmallCheckerboard,

    /// Large checkerboard - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_large_checkerboard.png">
    LargeCheckerboard,

    /// Outlined diamond grid - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_outlined_diamond_grid.png">
    OutlinedDiamondGrid,

    /// Solid diamond grid - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_solid_diamond_grid.png">
    SolidDiamondGrid,
}

impl ToString for ChartPatternFillType {
    fn to_string(&self) -> String {
        match self {
            ChartPatternFillType::Wave => "wave".to_string(),
            ChartPatternFillType::Weave => "weave".to_string(),
            ChartPatternFillType::Plaid => "plaid".to_string(),
            ChartPatternFillType::Divot => "divot".to_string(),
            ChartPatternFillType::Zigzag => "zigZag".to_string(),
            ChartPatternFillType::Sphere => "sphere".to_string(),
            ChartPatternFillType::Shingle => "shingle".to_string(),
            ChartPatternFillType::Trellis => "trellis".to_string(),
            ChartPatternFillType::SmallGrid => "smGrid".to_string(),
            ChartPatternFillType::LargeGrid => "lgGrid".to_string(),
            ChartPatternFillType::DottedGrid => "dotGrid".to_string(),
            ChartPatternFillType::DottedDiamond => "dotDmnd".to_string(),
            ChartPatternFillType::DiagonalBrick => "diagBrick".to_string(),
            ChartPatternFillType::LargeConfetti => "lgConfetti".to_string(),
            ChartPatternFillType::SmallConfetti => "smConfetti".to_string(),
            ChartPatternFillType::Dotted5Percent => "pct5".to_string(),
            ChartPatternFillType::Dotted10Percent => "pct10".to_string(),
            ChartPatternFillType::Dotted20Percent => "pct20".to_string(),
            ChartPatternFillType::Dotted25Percent => "pct25".to_string(),
            ChartPatternFillType::Dotted30Percent => "pct30".to_string(),
            ChartPatternFillType::Dotted40Percent => "pct40".to_string(),
            ChartPatternFillType::Dotted50Percent => "pct50".to_string(),
            ChartPatternFillType::Dotted60Percent => "pct60".to_string(),
            ChartPatternFillType::Dotted70Percent => "pct70".to_string(),
            ChartPatternFillType::Dotted75Percent => "pct75".to_string(),
            ChartPatternFillType::Dotted80Percent => "pct80".to_string(),
            ChartPatternFillType::Dotted90Percent => "pct90".to_string(),
            ChartPatternFillType::HorizontalBrick => "horzBrick".to_string(),
            ChartPatternFillType::SolidDiamondGrid => "solidDmnd".to_string(),
            ChartPatternFillType::SmallCheckerboard => "smCheck".to_string(),
            ChartPatternFillType::LargeCheckerboard => "lgCheck".to_string(),
            ChartPatternFillType::StripesBackslashes => "dashDnDiag".to_string(),
            ChartPatternFillType::VerticalStripesDark => "dkVert".to_string(),
            ChartPatternFillType::OutlinedDiamondGrid => "openDmnd".to_string(),
            ChartPatternFillType::VerticalStripesLight => "ltVert".to_string(),
            ChartPatternFillType::HorizontalStripesDark => "dkHorz".to_string(),
            ChartPatternFillType::StripesForwardSlashes => "dashUpDiag".to_string(),
            ChartPatternFillType::VerticalStripesNarrow => "narVert".to_string(),
            ChartPatternFillType::HorizontalStripesLight => "ltHorz".to_string(),
            ChartPatternFillType::HorizontalStripesNarrow => "narHorz".to_string(),
            ChartPatternFillType::DiagonalStripesDarkUpwards => "dkUpDiag".to_string(),
            ChartPatternFillType::DiagonalStripesWideUpwards => "wdUpDiag".to_string(),
            ChartPatternFillType::VerticalStripesAlternating => "dashVert".to_string(),
            ChartPatternFillType::DiagonalStripesLightUpwards => "ltUpDiag".to_string(),
            ChartPatternFillType::DiagonalStripesDarkDownwards => "dkDnDiag".to_string(),
            ChartPatternFillType::DiagonalStripesWideDownwards => "wdDnDiag".to_string(),
            ChartPatternFillType::HorizontalStripesAlternating => "dashHorz".to_string(),
            ChartPatternFillType::DiagonalStripesLightDownwards => "ltDnDiag".to_string(),
        }
    }
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use crate::chart::{Chart, ChartRange, ChartSeries, ChartType, XlsxError};
    use crate::test_functions::xml_to_vec;
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
