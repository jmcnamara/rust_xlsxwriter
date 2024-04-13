// chart - A module for creating the Excel Chart.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! # Working with Charts
//!
//! The sections below explain some of the options and features when working
//! with the [`Chart`] struct.
//!
//!
//!
//!
//! ## Creating a chart with `rust_xlsxwriter`
//!
//! The basis steps for creating a chart in `rust_xlsxwriter` are:
//!
//! - Create a new [`Chart`] object of the chart type you want.
//! - Add one or more series to the chart via [`Chart::add_series()`] to define
//!   the data you wish to plot.
//! - Add any formatting or additional feature you need.
//! - Add the chart to the worksheet via
//!   [`Worksheet::insert_chart()`](crate::Worksheet::insert_chart).
//!
//! These steps are shown in the example below which creates a minimal chart
//! that plots data in a worksheet. The program creates a new column [`Chart`],
//! adds a series via [`Chart::add_series()`] and add the value range the series
//! refers to via [`ChartSeries::set_values()`]:
//!
//! ```rust
//! # // This code is available in examples/app_chart_tutorial1.rs
//! #
//! use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!     let worksheet = workbook.add_worksheet();
//!     let bold = Format::new().set_bold();
//!
//!     // Add the worksheet data that the charts will refer to.
//!     let categories = ["Mon", "Tue", "Wed", "Thu", "Fri"];
//!     let values = [20, 40, 50, 30, 20];
//!
//!     worksheet.write_with_format(0, 0, "Day", &bold)?;
//!     worksheet.write_column(1, 0, categories)?;
//!
//!     worksheet.write_with_format(0, 1, "Sample", &bold)?;
//!     worksheet.write_column(1, 1, values)?;
//!
//!     // Create a new column chart.
//!     let mut chart = Chart::new(ChartType::Column);
//!
//!     // Configure the data series for the chart.
//!     chart.add_series().set_values("Sheet1!$B$2:$B$6");
//!
//!     // Add the chart to the worksheet.
//!     worksheet.insert_chart(0, 2, &chart)?;
//!
//!     workbook.save("chart_tutorial1.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! This produces a file like this:
//!
//! <img src="https://rustxlsxwriter.github.io/images/chart_tutorial1.png">
//!
//! To improve this a little we will add the category data that the value data
//! refers to. In this case it is the "Day" data in Column A. We extend the
//! program by adding the category values via [`ChartSeries::set_categories()`]:
//!
//! ```rust
//! # // This code is available in examples/app_chart_tutorial2.rs
//! #
//! # use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #     let worksheet = workbook.add_worksheet();
//! #     let bold = Format::new().set_bold();
//! #
//! #     // Add the worksheet data that the charts will refer to.
//! #     let categories = ["Mon", "Tue", "Wed", "Thu", "Fri"];
//! #     let values = [20, 40, 50, 30, 20];
//! #
//! #     worksheet.write_with_format(0, 0, "Day", &bold)?;
//! #     worksheet.write_column(1, 0, categories)?;
//! #
//! #     worksheet.write_with_format(0, 1, "Sample", &bold)?;
//! #     worksheet.write_column(1, 1, values)?;
//! #
//! #     // Create a new column chart.
//!     let mut chart = Chart::new(ChartType::Column);
//!
//!     // Configure the data series for the chart.
//!     chart
//!         .add_series()
//!         .set_categories("Sheet1!$A$2:$A$6")
//!         .set_values("Sheet1!$B$2:$B$6");
//!
//!     // Add the chart to the worksheet.
//!     worksheet.insert_chart(0, 2, &chart)?;
//! #
//! #     workbook.save("chart_tutorial2.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//! (Note, the section of the program where we write the worksheet data is the
//! same as the previous example and is omitted in this example. You can view it
//! in the source or the examples folder.)
//!
//! The updated file like this (the chart data range is highlighted):
//!
//! <img src="https://rustxlsxwriter.github.io/images/chart_tutorial2.png">
//!
//! The previous example works fine but it it contained some hard-coded cell
//! ranges like  `set_values("Sheet1!$B$2:$B$6")`. This is okay for simple
//! programs but if our example changed to have a different number of data items
//! then we would have to manually change the code to adjust for the new ranges.
//!
//! Fortunately, these hard-coded values are only used for the sake of the
//! example and `rust_xlsxwriter` provides APIs to handle these more
//! programmatically.
//!
//! In general `rust_xlsxwriter` always provides numeric APIs for any ranges in
//! Excel but when it makes ergonomic sense it also provides **secondary**
//! string based APIs. The previous example uses one of these secondary string
//! based APIs for demonstration purposes but for real applications you would
//! set the chart ranges using 5-tuple values like this:
//!
//! ```rust
//! # // This code is available in examples/app_chart_tutorial3.rs
//! #
//! # use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #     let worksheet = workbook.add_worksheet();
//! #     let bold = Format::new().set_bold();
//! #
//! #     // Add the worksheet data that the charts will refer to.
//! #     let categories = ["Mon", "Tue", "Wed", "Thu", "Fri"];
//! #     let values = [20, 40, 50, 30, 20];
//! #
//! #     worksheet.write_with_format(0, 0, "Day", &bold)?;
//! #     worksheet.write_column(1, 0, categories)?;
//! #
//! #     worksheet.write_with_format(0, 1, "Sample", &bold)?;
//! #     worksheet.write_column(1, 1, values)?;
//! #
//!     // Set some variables to define the chart range.
//!     let row_min = 1;
//!     let row_max = values.len() as u32;
//!     let col_cat = 0;
//!     let col_val = 1;
//!
//!     // Create a new column chart.
//!     let mut chart = Chart::new(ChartType::Column);
//!
//!     // Configure the data series for the chart.
//!     chart
//!         .add_series()
//!         .set_categories(("Sheet1", row_min, col_cat, row_max, col_cat))
//!         .set_values(("Sheet1", row_min, col_val, row_max, col_val));
//!
//!     // Add the chart to the worksheet.
//!     worksheet.insert_chart(0, 2, &chart)?;
//! #
//! #     workbook.save("chart_tutorial3.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! We could use hard coded row and column indexes for the chart ranges but the
//! variables make it more flexible and future proof. The output from this
//! example is exactly the same as the previous image above.
//!
//! Finally we can improve the output a bit more by inserting chart and axes
//! titles, by hiding the legend since it doesn't provide much information in
//! this case, and by shifting the chart a few pixels away from the data for
//! clarity via
//! [`Worksheet::insert_chart_with_offset()`](crate::Worksheet::insert_chart_with_offset).
//!
//! Here is the chart section with these changes:
//!
//!
//! ```rust
//! # // This code is available in examples/app_chart_tutorial4.rs
//! #
//! # use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #     let worksheet = workbook.add_worksheet();
//! #     let bold = Format::new().set_bold();
//! #
//! #     // Add the worksheet data that the charts will refer to.
//! #     let categories = ["Mon", "Tue", "Wed", "Thu", "Fri"];
//! #     let values = [20, 40, 50, 30, 20];
//! #
//! #     worksheet.write_with_format(0, 0, "Day", &bold)?;
//! #     worksheet.write_column(1, 0, categories)?;
//! #
//! #     worksheet.write_with_format(0, 1, "Sample", &bold)?;
//! #     worksheet.write_column(1, 1, values)?;
//! #
//! #     // Set some variables to define the chart range.
//! #     let row_min = 1;
//! #     let row_max = values.len() as u32;
//! #     let col_cat = 0;
//! #     let col_val = 1;
//! #
//!     // Create a new column chart.
//!     let mut chart = Chart::new(ChartType::Column);
//!
//!     // Configure the data series for the chart.
//!     chart
//!         .add_series()
//!         .set_categories(("Sheet1", row_min, col_cat, row_max, col_cat))
//!         .set_values(("Sheet1", row_min, col_val, row_max, col_val));
//!
//!     // Add a chart title and some axis labels.
//!     chart.title().set_name("Results of sample tests");
//!     chart.x_axis().set_name("Test day");
//!     chart.y_axis().set_name("Sample length (mm)");
//!
//!     // Turn off the chart legend.
//!     chart.legend().set_hidden();
//!
//!     // Add the chart to the worksheet.
//!     worksheet.insert_chart_with_offset(0, 2, &chart, 5, 5)?;
//! #
//! #     workbook.save("chart_tutorial4.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! This produces a file like this:
//!
//! <img src="https://rustxlsxwriter.github.io/images/chart_tutorial4.png">
//!
//! For more examples see chart examples in the[`Cookbook`](crate::cookbook).
//!
//!
//!
//!
//! ## Supported chart types
//!
//! The `rust_xlsxwriter` library supports the main original Excel chart types
//! such as:
//!
//! - Area
//! - Bar
//! - Column
//! - Doughnut
//! - Line
//! - Pie
//! - Radar
//! - Stock
//! - Scatter
//!
//! See [`ChartType`] for the full list and examples.
//!
//! Support for newer Excel chart types such as Treemap, Sunburst, Box and
//! Whisker, Statistical Histogram, Waterfall, Funnel, and Maps is not currently
//! planned since the underlying structure is substantially different from the
//! original chart types above.
//!
//!
//!
//! ## Chart formatting
//!
//! Excel uses a standard dialog for any chart elements that support formatting
//! such as data series, the plot area, the chart area, the legend or individual
//! points. It looks like this:
//!
//! <img src="https://rustxlsxwriter.github.io/images/chart_format_dialog.png">
//!
//! In `rust_xlsxwriter` the [`ChartFormat`] struct represents many of these
//! format options and just like Excel it offers a standard formatting interface
//! for a number of the chart elements.
//!
//! The [`ChartFormat`] struct is generally passed to the `set_format()` method
//! of a chart element. `ChartFormat` supports several child formatting structs
//! such as:
//!
//! - [`ChartSolidFill`] properties for solid fills.
//! - [`ChartPatternFill`] for pattern fills.
//! - [`ChartGradientFill`] for gradient fills.
//! - [`ChartLine`] for line/border properties.
//!
//! As a syntactic shortcut you can pass any of the child format structs to
//! `set_format()`. This allows you to create and use a [`ChartLine`] struct,
//! for example, without having to construct a parent [`ChartFormat`]. However,
//! if you mix more than one type of format you will need a parent
//! [`ChartFormat`] struct to group them.
//!
//!
//!
//!
//! ## Chart Value and Category Axes
//!
//! When working with charts it is important to understand how Excel
//! differentiates between a chart axis that is used for data series
//! "Categories" and a chart axis that is used for data series "Values".
//!
//! The majority of Excel charts types have a "Category" X-axis axis where each
//! of the values is evenly spaced and sequential. The category values can be
//! strings, like in the example below, or they can be numbers. When the
//! categories are numbers Excel treats them as if they were strings which is
//! why you can't set a minimum or maximum limit in Category axes.
//!
//! <img src="https://rustxlsxwriter.github.io/images/chart_axes01.png">
//!
//! <p style="text-align: center;"><b>Column chart with Category X-axis and Value Y-axis</b></p>
//!
//! The Y axis of Excel charts are generally a "Value" axis where points are
//! displayed according to their value. This type of axis does support minimum
//! and maximum limits.
//!
//! Scatter charts are an exception to the general rule since they have two
//! value axes in order to plot `(x, y)` data that doesn't necessarily lie on
//! fixed X-axis divisions, as shown in the chart below:
//!
//! <img src="https://rustxlsxwriter.github.io/images/chart_axes02.png">
//!
//! <p style="text-align: center;"><b>Scatter chart with Value X-axis and Value Y-axis</b></p>
//!
//! One other variation is a category style chart with date or time values for
//! the X-axis. This type of "Date" axis shares properties of Category and Value
//! axes. Date axes are used in Stock charts, or if you provide date/time data
//! and use the [`ChartAxis::set_date_axis()`] method.
//!
//! <img src="https://rustxlsxwriter.github.io/images/chart_axes03.png">
//!
//! <p style="text-align: center;"><b>Column chart with Date X-axis and Value Y-axis</b></p>
//!
//! Finally, one other variant is a Bar chart where the Category and Value axes
//! positions are reversed:
//!
//! <img src="https://rustxlsxwriter.github.io/images/chart_axes04.png">
//!
//! <p style="text-align: center;"><b>Bar chart with Category Y-axis and Value X-axis</b></p>
//!
//! In Excel category and values axes expose different properties as can be seen
//! in the dialogs for a each type of axis, shown below. Note that for Category
//! axes there is no minimum and maximum "Bounds" option:
//!
//! <img src="https://rustxlsxwriter.github.io/images/chart_axes05.png">
//!
//! <img src="https://rustxlsxwriter.github.io/images/chart_axes06.png">
//!
//! Due to Excel's distinction between axes type some `rust_xlsxwriter`
//! [`ChartAxis`] properties can be set for a value axis, some can be set for a
//! category axis and some properties can be set for both. For example `reverse`
//! can be set for either category or value axes while the `min` and `max`
//! properties can only be set for value axes (and date axes). The documentation
//! calls out the type of axis to which properties apply.
//!
//!
//!
//!
//! ## Future work
//!
//! Future additions to chart support in `rust_xlsxwriter` include:
//!
//! - Combined charts to allow x2/y2 axes.
//! - Chartsheets - Worksheets that only display a chart.
//! - Some chart element layout options.
//!
//! See the [Chart Roadmap] on the `rust_xlsxwriter` GitHub for more information.
//!
//! [Chart Roadmap]: https://github.com/jmcnamara/rust_xlsxwriter/issues/19
//!
#![warn(missing_docs)]

mod tests;

use regex::Regex;
use std::{fmt, mem};

use crate::{
    drawing::{DrawingObject, DrawingType},
    utility::{self, ToXmlBoolean},
    xmlwriter::XMLWriter,
    ColNum, Color, IntoColor, IntoExcelDateTime, ObjectMovement, RowNum, XlsxError, COL_MAX,
    ROW_MAX,
};

#[derive(Clone)]
/// The `Chart` struct is used to create an object to represent an chart that
/// can be inserted into a worksheet.
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
/// See also [Working with Charts](crate::chart) for a general introduction to
/// working with charts in `rust_xlsxwriter`.
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
pub struct Chart {
    pub(crate) id: u32,
    pub(crate) writer: XMLWriter,
    pub(crate) x_offset: u32,
    pub(crate) y_offset: u32,
    pub(crate) name: String,
    pub(crate) alt_text: String,
    pub(crate) object_movement: ObjectMovement,
    pub(crate) decorative: bool,
    pub(crate) drawing_type: DrawingType,
    pub(crate) series: Vec<ChartSeries>,
    pub(crate) default_label_position: ChartDataLabelPosition,
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
    pub(crate) combined_chart: Option<Box<Chart>>,
    grouping: ChartGrouping,
    show_empty_cells_as: Option<ChartEmptyCells>,
    show_hidden_data: bool,
    show_na_as_empty: bool,
    default_num_format: String,
    has_overlap: bool,
    overlap: i8,
    gap: u16,
    style: u8,
    hole_size: u8,
    rotation: u16,
    has_up_down_bars: bool,
    up_bar_format: ChartFormat,
    down_bar_format: ChartFormat,
    has_high_low_lines: bool,
    high_low_lines_format: ChartFormat,
    has_drop_lines: bool,
    drop_lines_format: ChartFormat,
    table: Option<ChartDataTable>,
    base_series_index: usize,
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
    /// series via [`chart.add_series()`](Chart::add_series) and set a value
    /// range for that series using
    /// [`series.set_values()`][ChartSeries::set_values]. See the example below.
    ///
    /// There are some shortcut versions of `new()` such as [`Chart::new_pie()`]
    /// that are more useful/succinct for charts that don't have subtypes.
    ///
    /// # Parameters
    ///
    /// `chart_type` - The chart type defined by [`ChartType`].
    ///
    /// # Examples
    ///
    /// A simple chart example using the `rust_xlsxwriter` library.
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
            name: String::new(),
            alt_text: String::new(),
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
            chart_area_format: ChartFormat::default(),
            plot_area_format: ChartFormat::default(),
            grouping: ChartGrouping::Standard,
            show_empty_cells_as: None,
            show_hidden_data: false,
            show_na_as_empty: false,
            default_num_format: "General".to_string(),
            has_overlap: false,
            overlap: 0,
            gap: 150,
            style: 2,
            hole_size: 50,
            rotation: 0,
            default_label_position: ChartDataLabelPosition::Default,
            has_up_down_bars: false,
            up_bar_format: ChartFormat::default(),
            down_bar_format: ChartFormat::default(),
            has_high_low_lines: false,
            high_low_lines_format: ChartFormat::default(),
            has_drop_lines: false,
            drop_lines_format: ChartFormat::default(),
            table: None,
            combined_chart: None,
            base_series_index: 0,
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

            ChartType::Stock => Self::initialize_stock_chart(chart),
        }
    }

    /// Create a new Area `Chart`.
    ///
    /// This is a syntactic shortcut for `Chart::new(ChartType::Area)` to
    /// create a default Area chart.
    ///
    /// See [`Chart::new()`] for further details.
    ///
    pub fn new_area() -> Chart {
        Self::new(ChartType::Area)
    }

    /// Create a new Bar `Chart`.
    ///
    /// This is a syntactic shortcut for `Chart::new(ChartType::Bar)` to
    /// create a default Bar chart.
    ///
    /// See [`Chart::new()`] for further details.
    ///
    pub fn new_bar() -> Chart {
        Self::new(ChartType::Bar)
    }

    /// Create a new Column `Chart`.
    ///
    /// This is a syntactic shortcut for `Chart::new(ChartType::Column)` to
    /// create a default Column chart.
    ///
    /// See [`Chart::new()`] for further details.
    ///
    pub fn new_column() -> Chart {
        Self::new(ChartType::Column)
    }

    /// Create a new Doughnut `Chart`.
    ///
    /// This is a syntactic shortcut for `Chart::new_doughnut()` to
    /// create a default Doughnut chart.
    ///
    /// See [`Chart::new()`] for further details.
    ///
    pub fn new_doughnut() -> Chart {
        Self::new(ChartType::Doughnut)
    }

    /// Create a new Line `Chart`.
    ///
    /// This is a syntactic shortcut for `Chart::new(ChartType::Line)` to
    /// create a default Line chart.
    ///
    /// See [`Chart::new()`] for further details.
    ///
    pub fn new_line() -> Chart {
        Self::new(ChartType::Line)
    }

    /// Create a new Pie `Chart`.
    ///
    /// This is a syntactic shortcut for `Chart::new(ChartType::Pie)` to
    /// create a default Pie chart.
    ///
    /// See [`Chart::new()`] for further details.
    ///
    pub fn new_pie() -> Chart {
        Self::new(ChartType::Pie)
    }

    /// Create a new Radar `Chart`.
    ///
    /// This is a syntactic shortcut for `Chart::new(ChartType::Radar)` to
    /// create a default Radar chart.
    ///
    /// See [`Chart::new()`] for further details.
    ///
    pub fn new_radar() -> Chart {
        Self::new(ChartType::Radar)
    }

    /// Create a new Scatter `Chart`.
    ///
    /// This is a syntactic shortcut for `Chart::new(ChartType::Scatter)` to
    /// create a default Scatter chart.
    ///
    /// See [`Chart::new()`] for further details.
    ///
    pub fn new_scatter() -> Chart {
        Self::new(ChartType::Scatter)
    }

    /// Create a new Stock `Chart`.
    ///
    /// This is a syntactic shortcut for `Chart::new(ChartType::Stock)` to
    /// create a default Stock chart.
    ///
    /// See [`Chart::new()`] for further details.
    ///
    pub fn new_stock() -> Chart {
        Self::new(ChartType::Stock)
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_simple.png">
    ///
    pub fn add_series(&mut self) -> &mut ChartSeries {
        let mut series = ChartSeries::new();

        // The default Scatter chart has a hidden line with a standard width.
        if self.chart_type == ChartType::Scatter {
            series.set_format(
                ChartFormat::new().set_line(ChartLine::new().set_width(2.25).set_hidden(true)),
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
    /// # Parameters
    ///
    /// `series` - a [`ChartSeries`] instance.
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
    /// #
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
                ChartFormat::new().set_line(ChartLine::new().set_width(2.25).set_hidden(true)),
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_name.png">
    ///
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_legend.png">
    ///
    pub fn legend(&mut self) -> &mut ChartLegend {
        &mut self.legend
    }

    /// Create a combination chart with a secondary chart.
    ///
    /// TODO explain chart `combine()`.
    ///
    ///
    pub fn combine(&mut self, chart: &Chart) -> &mut Chart {
        self.combined_chart = Some(Box::new(chart.clone()));

        self
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
    /// # Parameters
    ///
    /// * `style` - A integer value in the range 1-48.
    ///
    /// # Examples
    ///
    /// An example showing all 48 default chart styles available in Excel 2007
    /// using `rust_xlsxwriter`.
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
            eprintln!("Style id '{style}' outside Excel range: 1 <= style <= 48.");
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
    /// - [`ChartFormat::set_solid_fill()`]: Set the [`ChartSolidFill`] properties.
    /// - [`ChartFormat::set_pattern_fill()`]: Set the [`ChartPatternFill`] properties.
    /// - [`ChartFormat::set_gradient_fill()`]: Set the [`ChartGradientFill`] properties.
    /// - [`ChartFormat::set_no_fill()`]: Turn off the fill for the chart object.
    /// - [`ChartFormat::set_line()`]: Set the [`ChartLine`] properties.
    /// - [`ChartFormat::set_border()`]: Set the [`ChartBorder`] properties.
    ///   A synonym for [`ChartLine`] depending on context.
    /// - [`ChartFormat::set_no_line()`]: Turn off the line for the chart object.
    /// - [`ChartFormat::set_no_border()`]: Turn off the border for the chart object.
    ///
    /// # Parameters
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
    /// - [`ChartFormat::set_solid_fill()`]: Set the [`ChartSolidFill`] properties.
    /// - [`ChartFormat::set_pattern_fill()`]: Set the [`ChartPatternFill`] properties.
    /// - [`ChartFormat::set_gradient_fill()`]: Set the [`ChartGradientFill`] properties.
    /// - [`ChartFormat::set_no_fill()`]: Turn off the fill for the chart object.
    /// - [`ChartFormat::set_line()`]: Set the [`ChartLine`] properties.
    /// - [`ChartFormat::set_border()`]: Set the [`ChartBorder`] properties.
    ///   A synonym for [`ChartLine`] depending on context.
    /// - [`ChartFormat::set_no_line()`]: Turn off the line for the chart object.
    /// - [`ChartFormat::set_no_border()`]: Turn off the border for the chart object.
    ///
    /// # Parameters
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
    /// # Parameters
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
    /// # use rust_xlsxwriter::{Chart, Workbook, XlsxError};
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
    ///     let mut chart = Chart::new_pie();
    ///
    ///     // Add a data series with formatting.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Set the rotation of the chart.
    ///     chart.set_rotation(270);
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// # Parameters
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
    /// # use rust_xlsxwriter::{Chart, Workbook, XlsxError};
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
    ///     let mut chart = Chart::new_doughnut();
    ///
    ///     // Add a data series with formatting.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Set the home size of the chart.
    ///     chart.set_hole_size(80);
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_set_hole_size.png">
    ///
    pub fn set_hole_size(&mut self, hole_size: u8) -> &mut Chart {
        if (0..=90).contains(&hole_size) {
            self.hole_size = hole_size;
        }
        self
    }

    /// Set Up-Down bar indicators for a Line chart.
    ///
    /// Set Up-Down bar indicator to indicate change between two or more series.
    /// In Excel these can only be added to Line and Stock charts.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// An example of setting up-down bars for a chart.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_up_down_bars.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     let data = [[5, 10, 15], [4, 9, 13], [3, 8, 10], [2, 7, 6], [1, 6, 4]];
    /// #
    /// #     for (row_num, row_data) in data.iter().enumerate() {
    /// #         for (col_num, col_data) in row_data.iter().enumerate() {
    /// #             worksheet.write_number(row_num as u32, col_num as u16, *col_data)?;
    /// #         }
    /// #     }
    /// #
    /// #     // Create the chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 0, 4, 0))
    ///         .set_values(("Sheet1", 0, 1, 4, 1));
    ///
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 0, 4, 0))
    ///         .set_values(("Sheet1", 0, 2, 4, 2));
    ///
    ///     // Set the up-down bars.
    ///     chart.set_up_down_bars(true);
    ///
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
    /// src="https://rustxlsxwriter.github.io/images/chart_set_up_down_bars.png">
    ///
    pub fn set_up_down_bars(&mut self, enable: bool) -> &mut Chart {
        self.has_up_down_bars = enable;
        self
    }

    /// Set the formatting properties for Line chart up bars.
    ///
    /// Set the formatting properties for Line chart positive "Up" bars via a
    /// [`ChartFormat`] object or a sub struct that implements
    /// [`IntoChartFormat`].
    ///
    /// See [`ChartFormat`] for the format properties that can be set.
    ///
    /// # Parameters
    ///
    /// `format`: A [`ChartFormat`] struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// [`IntoChartFormat`] for details.
    ///
    ///
    /// # Examples
    ///
    /// An example of setting up-down bars for a chart, with formatting.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_up_down_bars_format.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFormat, ChartSolidFill, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     let data = [[5, 10, 15], [4, 9, 13], [3, 8, 10], [2, 7, 6], [1, 6, 4]];
    /// #
    /// #     for (row_num, row_data) in data.iter().enumerate() {
    /// #         for (col_num, col_data) in row_data.iter().enumerate() {
    /// #             worksheet.write_number(row_num as u32, col_num as u16, *col_data)?;
    /// #         }
    /// #     }
    /// #
    /// #     // Create the chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 0, 4, 0))
    ///         .set_values(("Sheet1", 0, 1, 4, 1));
    ///
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 0, 4, 0))
    ///         .set_values(("Sheet1", 0, 2, 4, 2));
    ///
    ///     // Set the up-down bars.
    ///     chart
    ///         .set_up_down_bars(true)
    ///         .set_up_bar_format(
    ///             ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#00B050")),
    ///         )
    ///         .set_down_bar_format(
    ///             ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FF0000")),
    ///         );
    ///
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
    /// src="https://rustxlsxwriter.github.io/images/chart_set_up_down_bars_format.png">
    ///
    pub fn set_up_bar_format<T>(&mut self, format: T) -> &mut Chart
    where
        T: IntoChartFormat,
    {
        self.has_up_down_bars = true;
        self.up_bar_format = format.new_chart_format();
        self
    }

    /// Set the formatting properties for Line chart down bars.
    ///
    /// Set the formatting for negative "Down" bars on an "Up-Down" chart
    /// element. See the documentation for [`Chart::set_up_bar_format()`].
    ///
    pub fn set_down_bar_format<T>(&mut self, format: T) -> &mut Chart
    where
        T: IntoChartFormat,
    {
        self.has_up_down_bars = true;
        self.down_bar_format = format.new_chart_format();
        self
    }

    /// Set High-Low lines for a Line chart.
    ///
    /// Set High-Low lines for a Line chart to indicate the high and low values
    /// between two or more series. In Excel these can only be added to Line and
    /// Stock charts.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// An example of setting high-low lines for a chart.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_high_low_lines.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     let data = [[5, 10, 15], [4, 9, 13], [3, 8, 10], [2, 7, 6], [1, 6, 4]];
    /// #
    /// #     for (row_num, row_data) in data.iter().enumerate() {
    /// #         for (col_num, col_data) in row_data.iter().enumerate() {
    /// #             worksheet.write_number(row_num as u32, col_num as u16, *col_data)?;
    /// #         }
    /// #     }
    /// #
    /// #     // Create the chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 0, 4, 0))
    ///         .set_values(("Sheet1", 0, 1, 4, 1));
    ///
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 0, 4, 0))
    ///         .set_values(("Sheet1", 0, 2, 4, 2));
    ///
    ///     // Set the high_low lines.
    ///     chart.set_high_low_lines(true);
    ///
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
    /// src="https://rustxlsxwriter.github.io/images/chart_set_high_low_lines.png">
    ///
    pub fn set_high_low_lines(&mut self, enable: bool) -> &mut Chart {
        self.has_high_low_lines = enable;
        self
    }

    /// Set the formatting properties for Line chart High-Low lines.
    ///
    /// Set the formatting properties for line chart high-low lines via a
    /// [`ChartFormat`] object or a sub struct that implements
    /// [`IntoChartFormat`]. In general you will only need to use a
    /// [`ChartLine`] to define the line format properties.
    ///
    /// # Parameters
    ///
    /// * `format`: A [`ChartFormat`] struct reference or a sub struct that will
    ///   convert into a `ChartFormat` instance. See the docs for
    ///   [`IntoChartFormat`] for details.
    ///
    /// # Examples
    ///
    /// An example of setting high-low lines for a chart, with formatting.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_high_low_lines_format.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartLine, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     let data = [[5, 10, 15], [4, 9, 13], [3, 8, 10], [2, 7, 6], [1, 6, 4]];
    /// #
    /// #     for (row_num, row_data) in data.iter().enumerate() {
    /// #         for (col_num, col_data) in row_data.iter().enumerate() {
    /// #             worksheet.write_number(row_num as u32, col_num as u16, *col_data)?;
    /// #         }
    /// #     }
    /// #
    /// #     // Create the chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 0, 4, 0))
    ///         .set_values(("Sheet1", 0, 1, 4, 1));
    ///
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 0, 4, 0))
    ///         .set_values(("Sheet1", 0, 2, 4, 2));
    ///
    ///     // Set the high_low lines.
    ///     chart
    ///         .set_high_low_lines(true)
    ///         .set_high_low_lines_format(ChartLine::new().set_color("#FF0000"));
    ///
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
    /// src="https://rustxlsxwriter.github.io/images/chart_set_high_low_lines_format.png">
    ///
    pub fn set_high_low_lines_format<T>(&mut self, format: T) -> &mut Chart
    where
        T: IntoChartFormat,
    {
        self.has_high_low_lines = true;
        self.high_low_lines_format = format.new_chart_format();
        self
    }

    /// Set drop lines for a chart.
    ///
    /// Set drop lines for a chart between the maximum value and the associated
    /// category value. In Excel these can only be added to Line, Area and Stock
    /// charts.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    /// # Examples
    ///
    /// An example of setting drop lines for a chart.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_drop_lines.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     let data = [[5, 10, 15], [4, 9, 13], [3, 8, 10], [2, 7, 6], [1, 6, 4]];
    /// #
    /// #     for (row_num, row_data) in data.iter().enumerate() {
    /// #         for (col_num, col_data) in row_data.iter().enumerate() {
    /// #             worksheet.write_number(row_num as u32, col_num as u16, *col_data)?;
    /// #         }
    /// #     }
    /// #
    /// #     // Create the chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 0, 4, 0))
    ///         .set_values(("Sheet1", 0, 1, 4, 1));
    ///
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 0, 4, 0))
    ///         .set_values(("Sheet1", 0, 2, 4, 2));
    ///
    ///     // Set the drop lines.
    ///     chart.set_drop_lines(true);
    ///
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
    /// src="https://rustxlsxwriter.github.io/images/chart_set_drop_lines.png">
    ///
    pub fn set_drop_lines(&mut self, enable: bool) -> &mut Chart {
        self.has_drop_lines = enable;
        self
    }

    /// Set the formatting properties for a chart drop lines.
    ///
    /// Set the formatting properties for a chart drop lines via a
    /// [`ChartFormat`] object or a sub struct that implements
    /// [`IntoChartFormat`]. In general you will only need to use a
    /// [`ChartLine`] to define the line format properties.
    ///
    /// # Parameters
    ///
    /// * `format`: A [`ChartFormat`] struct reference or a sub struct that will
    ///   convert into a `ChartFormat` instance. See the docs for
    ///   [`IntoChartFormat`] for details.
    ///
    /// # Examples
    ///
    /// An example of setting drop lines for a chart, with formatting.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_drop_lines_format.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartLine, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     let data = [[5, 10, 15], [4, 9, 13], [3, 8, 10], [2, 7, 6], [1, 6, 4]];
    /// #
    /// #     for (row_num, row_data) in data.iter().enumerate() {
    /// #         for (col_num, col_data) in row_data.iter().enumerate() {
    /// #             worksheet.write_number(row_num as u32, col_num as u16, *col_data)?;
    /// #         }
    /// #     }
    /// #
    /// #     // Create the chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 0, 4, 0))
    ///         .set_values(("Sheet1", 0, 1, 4, 1));
    ///
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 0, 4, 0))
    ///         .set_values(("Sheet1", 0, 2, 4, 2));
    ///
    ///     // Set the drop lines.
    ///     chart
    ///         .set_drop_lines(true)
    ///         .set_drop_lines_format(ChartLine::new().set_color("#FF0000"));
    ///
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
    /// src="https://rustxlsxwriter.github.io/images/chart_set_drop_lines_format.png">
    ///
    pub fn set_drop_lines_format<T>(&mut self, format: T) -> &mut Chart
    where
        T: IntoChartFormat,
    {
        self.has_drop_lines = true;
        self.drop_lines_format = format.new_chart_format();
        self
    }

    /// Set a data table for a chart.
    ///
    /// A chart data table in Excel is an additional table below a chart that
    /// shows the plotted data in tabular form.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_data_table.png">
    ///
    /// The chart data table has the following default properties which can be
    /// set via the [`ChartDataTable`] struct.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_data_table_options.png">
    ///
    /// # Parameters
    ///
    /// * `table`: A [`ChartDataTable`] reference.
    ///
    /// # Examples
    ///
    /// An example of adding a data table to a chart.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_data_table.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartDataTable, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     let data = [[1, 2, 3], [2, 4, 6], [3, 6, 9], [4, 8, 12], [5, 10, 15]];
    /// #     for (row_num, row_data) in data.iter().enumerate() {
    /// #         for (col_num, col_data) in row_data.iter().enumerate() {
    /// #             worksheet.write_number(row_num as u32, col_num as u16, *col_data)?;
    /// #         }
    /// #     }
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new_column();
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///     chart.add_series().set_values("Sheet1!$B$1:$B$5");
    ///     chart.add_series().set_values("Sheet1!$C$1:$C$5");
    ///
    ///     // Add a default data table to the chart.
    ///     let table = ChartDataTable::default();
    ///     chart.set_data_table(&table);
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
    /// src="https://rustxlsxwriter.github.io/images/chart_set_data_table.png">
    ///
    pub fn set_data_table(&mut self, table: &ChartDataTable) -> &mut Chart {
        self.table = Some(table.clone());
        self
    }

    /// Set the width of the chart.
    ///
    /// The default width of an Excel chart is 480 pixels. The `set_width()`
    /// method allows you to set it to some other non-zero size.
    ///
    /// # Parameters
    ///
    /// * `width` - The chart width in pixels.
    ///
    /// # Examples
    ///
    /// A simple chart example using the `rust_xlsxwriter` library.
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_set_width.png">
    ///
    pub fn set_width(&mut self, width: u32) -> &mut Chart {
        if width == 0 {
            return self;
        }

        self.width = f64::from(width);
        self
    }

    /// Set the height of the chart.
    ///
    /// The default height of an Excel chart is 480 pixels. The `set_height()`
    /// method allows you to set it to some other non-zero size. See the example
    /// above.
    ///
    /// # Parameters
    ///
    /// * `height` - The chart height in pixels.
    ///
    pub fn set_height(&mut self, height: u32) -> &mut Chart {
        if height == 0 {
            return self;
        }

        self.height = f64::from(height);
        self
    }

    /// Set the height scale for the chart.
    ///
    /// Set the height scale for the chart relative to 1.0/100%. This is a
    /// syntactic alternative to [`chart.set_height()`](Chart::set_height).
    ///
    /// # Parameters
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
    /// # Parameters
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

    /// Set a user defined name for a chart.
    ///
    /// By default Excel names charts as "Chart 1", "Chart 2", etc. This name
    /// shows up in the formula bar and can be used to find or reference a
    /// chart.
    ///
    /// The [`Chart::set_name()`] method allows you to give the chart a user
    /// defined name.
    ///
    /// # Parameters
    ///
    /// * `name` - A user defined name for the chart.
    ///
    pub fn set_name(&mut self, name: impl Into<String>) -> &mut Chart {
        self.name = name.into();
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
    /// # Parameters
    ///
    /// * `alt_text` - The alt text string to add to the chart.
    ///
    pub fn set_alt_text(&mut self, alt_text: impl Into<String>) -> &mut Chart {
        self.alt_text = alt_text.into();
        self
    }

    /// Mark a chart as decorative.
    ///
    /// Charts don't always need an alt text description. Some charts may contain
    /// little or no useful visual information. Such charts can be marked as
    /// "decorative" so that screen readers can inform the users that they don't
    /// contain important information.
    ///
    /// # Parameters
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
    /// # Parameters
    ///
    /// `option` - A [`ObjectMovement`] enum value.
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
    ///   chart parameter is incorrect or a chart is configured incorrectly.
    /// * [`XlsxError::SheetnameCannotBeBlank`] - Worksheet name in chart range
    ///   cannot be blank.
    /// * [`XlsxError::SheetnameLengthExceeded`] - Worksheet name in chart range
    ///   exceeds Excel's limit of 31 characters.
    /// * [`XlsxError::SheetnameContainsInvalidCharacter`] - Worksheet name in
    ///   chart range cannot contain invalid characters: `[ ] : * ? / \`
    /// * [`XlsxError::SheetnameStartsOrEndsWithApostrophe`] - Worksheet name in
    ///   chart range cannot start or end with an apostrophe.
    ///
    pub fn validate(&mut self) -> Result<&mut Chart, XlsxError> {
        // Check for chart without series.
        if self.series.is_empty() {
            return Err(XlsxError::ChartError(
                "Chart must contain at least one series".to_string(),
            ));
        }

        for series in &self.series {
            // Check for a series without a values range.
            if !series.value_range.has_data() {
                return Err(XlsxError::ChartError(
                    "Chart series must contain a 'values' range".to_string(),
                ));
            }

            // Check for scatter charts without category ranges. It is optional
            // for all other types.
            if self.chart_group_type == ChartType::Scatter && !series.category_range.has_data() {
                return Err(XlsxError::ChartError(
                    "Scatter style charts must contain a 'categories' range".to_string(),
                ));
            }

            // Validate the series values range.
            series.value_range.validate()?;

            // Validate the series category range.
            if series.category_range.has_data() {
                series.category_range.validate()?;
            }

            // Validate Polynomial trendline range.
            if let ChartTrendlineType::Polynomial(order) = series.trendline.trend_type {
                if !(2..6).contains(&order) {
                    return Err(XlsxError::ChartError(
                        "Chart series Polynomial trendline order must be in the Excel range 2-6"
                            .to_string(),
                    ));
                }
            }

            // Validate Moving Average trendline range.
            if let ChartTrendlineType::MovingAverage(period) = series.trendline.trend_type {
                if !(2..4).contains(&period) {
                    return Err(XlsxError::ChartError(
                        "Chart series Moving Average trendline period must be in the Excel range 2-4"
                            .to_string(),
                    ));
                }
            }
        }

        Ok(self)
    }

    /// Set the option for displaying empty cells in a chart.
    ///
    /// The options are:
    ///
    /// * [`ChartEmptyCells::Gaps`]: Show empty cells in the chart as gaps. The
    ///   default.
    /// * [`ChartEmptyCells::Zero`]: Show empty cells in the chart as zeroes.
    /// * [`ChartEmptyCells::Connected`]: Show empty cells in the chart
    ///   connected by a line to the previous point.
    ///
    /// # Parameters
    ///
    /// `option` - A [`ChartEmptyCells`] enum value.
    ///
    pub fn show_empty_cells_as(&mut self, option: ChartEmptyCells) -> &mut Chart {
        self.show_empty_cells_as = Some(option);

        self
    }

    /// Display #N/A on charts as blank/empty cells.
    ///
    pub fn show_na_as_empty_cell(&mut self) -> &mut Chart {
        self.show_na_as_empty = true;

        self
    }

    /// Display data on charts from hidden rows or columns.
    ///
    pub fn show_hidden_data(&mut self) -> &mut Chart {
        self.show_hidden_data = true;

        self
    }

    /// Set default values for the chart axis ids.
    ///
    /// This is mainly used to ensure that the axis ids used in testing match
    /// the semi-randomized values in the target Excel file.
    ///
    /// # Parameters
    ///
    /// `axis_id1` - X-axis id.
    /// `axis_id2` - Y-axis id.
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

    // Check for any legend entries that have been hidden/deleted via the
    // ChartSeries::delete_from_legend() and
    // ChartTrendline::delete_from_legend() methods. These can in turn be
    // overridden by the `ChartLegend::delete_entries()` method, which is
    // checked for first.
    fn deleted_legend_entries(&self) -> Vec<usize> {
        // Use the user supplied entries, if available.
        if !self.legend.deleted_entries.is_empty() {
            return self.legend.deleted_entries.clone();
        }

        let mut deleted_entries = vec![];
        let mut index = 0;

        // Check for deleted series in legend.
        for series in &self.series {
            if series.delete_from_legend {
                deleted_entries.push(index);
            }

            index += 1;
        }

        // Check for deleted trendlines in legend. These are indexed after the
        // series they belong to.
        for series in &self.series {
            if series.trendline.trend_type != ChartTrendlineType::None {
                if series.trendline.delete_from_legend {
                    deleted_entries.push(index);
                }

                index += 1;
            }
        }

        deleted_entries
    }

    // -----------------------------------------------------------------------
    // Chart specific methods.
    // -----------------------------------------------------------------------

    // Initialize area charts.
    fn initialize_area_chart(mut self) -> Chart {
        self.x_axis.axis_type = ChartAxisType::Category;
        self.x_axis.axis_position = ChartAxisPosition::Bottom;
        self.x_axis.position_between_ticks = false;

        self.y_axis.axis_type = ChartAxisType::Value;
        self.y_axis.axis_position = ChartAxisPosition::Left;
        self.y_axis.title.is_horizontal = true;
        self.y_axis.major_gridlines = true;

        self.chart_group_type = ChartType::Area;

        if self.chart_type == ChartType::Area {
            self.grouping = ChartGrouping::Standard;
        } else if self.chart_type == ChartType::AreaStacked {
            self.grouping = ChartGrouping::Stacked;
        } else if self.chart_type == ChartType::AreaPercentStacked {
            self.grouping = ChartGrouping::PercentStacked;
            self.default_num_format = "0%".to_string();
        }

        self.default_label_position = ChartDataLabelPosition::Center;

        self
    }

    // Initialize bar charts. Bar chart category/value axes are reversed in
    // comparison to other charts. Some of the defaults reflect this.
    fn initialize_bar_chart(mut self) -> Chart {
        self.x_axis.axis_type = ChartAxisType::Value;
        self.x_axis.axis_position = ChartAxisPosition::Bottom;
        self.x_axis.major_gridlines = true;

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

        self.default_label_position = ChartDataLabelPosition::OutsideEnd;

        self
    }

    // Initialize column charts.
    fn initialize_column_chart(mut self) -> Chart {
        self.x_axis.axis_type = ChartAxisType::Category;
        self.x_axis.axis_position = ChartAxisPosition::Bottom;

        self.y_axis.axis_type = ChartAxisType::Value;
        self.y_axis.axis_position = ChartAxisPosition::Left;
        self.y_axis.major_gridlines = true;

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

        self.default_label_position = ChartDataLabelPosition::OutsideEnd;

        self
    }

    // Initialize doughnut charts.
    fn initialize_doughnut_chart(mut self) -> Chart {
        self.chart_group_type = ChartType::Doughnut;

        self.default_label_position = ChartDataLabelPosition::BestFit;

        self
    }

    // Initialize line charts.
    fn initialize_line_chart(mut self) -> Chart {
        self.x_axis.axis_type = ChartAxisType::Category;
        self.x_axis.axis_position = ChartAxisPosition::Bottom;

        self.y_axis.axis_type = ChartAxisType::Value;
        self.y_axis.axis_position = ChartAxisPosition::Left;
        self.y_axis.title.is_horizontal = true;
        self.y_axis.major_gridlines = true;

        self.chart_group_type = ChartType::Line;

        if self.chart_type == ChartType::Line {
            self.grouping = ChartGrouping::Standard;
        } else if self.chart_type == ChartType::LineStacked {
            self.grouping = ChartGrouping::Stacked;
        } else if self.chart_type == ChartType::LinePercentStacked {
            self.grouping = ChartGrouping::PercentStacked;
            self.default_num_format = "0%".to_string();
        }

        self.default_label_position = ChartDataLabelPosition::Right;

        self
    }

    // Initialize pie charts.
    fn initialize_pie_chart(mut self) -> Chart {
        self.chart_group_type = ChartType::Pie;

        self.default_label_position = ChartDataLabelPosition::BestFit;

        self
    }

    // Initialize radar charts.
    fn initialize_radar_chart(mut self) -> Chart {
        self.x_axis.axis_type = ChartAxisType::Category;
        self.x_axis.axis_position = ChartAxisPosition::Bottom;
        self.x_axis.major_gridlines = true;

        self.y_axis.axis_type = ChartAxisType::Value;
        self.y_axis.axis_position = ChartAxisPosition::Left;
        self.y_axis.major_gridlines = true;
        self.y_axis.major_tick_type = Some(ChartAxisTickType::Cross);

        self.chart_group_type = ChartType::Radar;

        self.default_label_position = ChartDataLabelPosition::Center;

        self
    }

    // Initialize scatter charts.
    fn initialize_scatter_chart(mut self) -> Chart {
        self.x_axis.axis_type = ChartAxisType::Value;
        self.x_axis.axis_position = ChartAxisPosition::Bottom;
        self.x_axis.position_between_ticks = false;

        self.y_axis.axis_type = ChartAxisType::Value;
        self.y_axis.axis_position = ChartAxisPosition::Left;
        self.y_axis.position_between_ticks = false;
        self.y_axis.title.is_horizontal = true;
        self.y_axis.major_gridlines = true;

        self.chart_group_type = ChartType::Scatter;

        self.default_label_position = ChartDataLabelPosition::Right;

        self
    }

    // Initialize stock charts.
    fn initialize_stock_chart(mut self) -> Chart {
        self.x_axis.axis_type = ChartAxisType::Date;
        self.x_axis.axis_position = ChartAxisPosition::Bottom;
        self.x_axis.automatic = true;

        self.y_axis.axis_type = ChartAxisType::Value;
        self.y_axis.axis_position = ChartAxisPosition::Left;
        self.y_axis.title.is_horizontal = true;
        self.y_axis.major_gridlines = true;

        self.chart_group_type = ChartType::Stock;
        self.default_label_position = ChartDataLabelPosition::Right;

        self
    }

    // Write the <c:areaChart> element for Column charts.
    fn write_area_chart(&mut self) {
        self.writer.xml_start_tag_only("c:areaChart");

        // Write the c:grouping element.
        self.write_grouping();

        // Write the c:ser elements.
        self.write_series();

        if self.has_drop_lines {
            // Write the c:dropLines element.
            self.write_drop_lines();
        }

        // Write the c:axId elements.
        self.write_ax_ids();

        self.writer.xml_end_tag("c:areaChart");
    }

    // Write the <c:barChart> element for Bar charts.
    fn write_bar_chart(&mut self) {
        self.writer.xml_start_tag_only("c:barChart");

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
        self.writer.xml_start_tag_only("c:barChart");

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
        self.writer.xml_start_tag_only("c:doughnutChart");

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
        self.writer.xml_start_tag_only("c:lineChart");

        // Write the c:grouping element.
        self.write_grouping();

        // Write the c:ser elements.
        self.write_series();

        if self.has_drop_lines {
            // Write the c:dropLines element.
            self.write_drop_lines();
        }

        if self.has_high_low_lines {
            // Write the c:hiLowLines element.
            self.write_hi_low_lines();
        }

        // Write the c:upDownBars element.
        if self.has_up_down_bars {
            self.write_up_down_bars();
        }

        // Write the c:marker element.
        self.write_marker_value();

        // Write the c:axId elements.
        self.write_ax_ids();

        self.writer.xml_end_tag("c:lineChart");
    }

    // Write the <c:pieChart> element for Column charts.
    fn write_pie_chart(&mut self) {
        self.writer.xml_start_tag_only("c:pieChart");

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
        self.writer.xml_start_tag_only("c:radarChart");

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
        self.writer.xml_start_tag_only("c:scatterChart");

        // Write the c:scatterStyle element.
        self.write_scatter_style();

        // Write the c:ser elements.
        self.write_scatter_series();

        // Write the c:axId elements.
        self.write_ax_ids();

        self.writer.xml_end_tag("c:scatterChart");
    }

    // Write the <c:stockChart>element.
    fn write_stock_chart(&mut self) {
        self.writer.xml_start_tag_only("c:stockChart");

        // Write the c:ser elements.
        self.write_series();

        if self.has_drop_lines {
            // Write the c:dropLines element.
            self.write_drop_lines();
        }

        if self.has_high_low_lines {
            // Write the c:hiLowLines element.
            self.write_hi_low_lines();
        }

        // Write the c:upDownBars element.
        if self.has_up_down_bars {
            self.write_up_down_bars();
        }

        // Write the c:axId elements.
        self.write_ax_ids();

        self.writer.xml_end_tag("c:stockChart");
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
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
        let attributes = [
            (
                "xmlns:c",
                "http://schemas.openxmlformats.org/drawingml/2006/chart",
            ),
            (
                "xmlns:a",
                "http://schemas.openxmlformats.org/drawingml/2006/main",
            ),
            (
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            ),
        ];

        self.writer.xml_start_tag("c:chartSpace", &attributes);
    }

    // Write the <c:lang> element.
    fn write_lang(&mut self) {
        let attributes = [("val", "en-US")];

        self.writer.xml_empty_tag("c:lang", &attributes);
    }

    // Write the <c:chart> element.
    fn write_chart(&mut self) {
        self.writer.xml_start_tag_only("c:chart");

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
        if !self.show_hidden_data {
            self.write_plot_vis_only();
        }

        // Write the c:dispBlanksAs element.
        self.write_disp_blanks_as();

        // Write the dispNaAsBlank element.
        if self.show_na_as_empty {
            self.write_disp_na_as_blank();
        }

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
        self.writer.xml_start_tag_only("c:plotArea");

        // Write the c:layout element.
        self.write_layout();

        // Write the <c:xxxChart> element for each chart type.
        self.write_chart_type();

        // Write the combined chart.
        if let Some(combined_chart) = &mut self.combined_chart {
            combined_chart.axis_ids = self.axis_ids;
            combined_chart.base_series_index = self.series.len();

            mem::swap(&mut combined_chart.writer, &mut self.writer);
            combined_chart.write_chart_type();
            mem::swap(&mut combined_chart.writer, &mut self.writer);
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
                if self.x_axis.axis_type == ChartAxisType::Date {
                    // Write the c:dateAx element.
                    self.write_date_ax();
                } else {
                    // Write the c:catAx element.
                    self.write_cat_ax();
                }

                // Write the c:valAx element.
                self.write_val_ax();
            }
        }

        // Reset the X and Y axes for Bar charts.
        if self.chart_group_type == ChartType::Bar {
            std::mem::swap(&mut self.x_axis, &mut self.y_axis);
        }

        // Write the c:dTable element.
        if let Some(table) = &self.table {
            self.write_data_table(&table.clone());
        }

        // Write the c:spPr element.
        self.write_sp_pr(&self.plot_area_format.clone());

        self.writer.xml_end_tag("c:plotArea");
    }

    // Write the <c:xxxChart> element.
    fn write_chart_type(&mut self) {
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

            ChartType::Stock => {
                self.write_stock_chart();
            }
        }
    }

    // Write the <c:layout> element.
    fn write_layout(&mut self) {
        self.writer.xml_empty_tag_only("c:layout");
    }

    // Write the <c:barDir> element.
    fn write_bar_dir(&mut self, direction: &str) {
        let attributes = [("val", direction.to_string())];

        self.writer.xml_empty_tag("c:barDir", &attributes);
    }

    // Write the <c:grouping> element.
    fn write_grouping(&mut self) {
        let attributes = [("val", self.grouping.to_string())];

        self.writer.xml_empty_tag("c:grouping", &attributes);
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

        self.writer.xml_empty_tag("c:scatterStyle", &attributes);
    }

    // Write the <c:ser> element.
    fn write_series(&mut self) {
        for (index, series) in self.series.clone().iter_mut().enumerate() {
            let max_points = series.value_range.number_of_points();

            self.writer.xml_start_tag_only("c:ser");

            // Copy a series overlap to the parent chart.
            if series.overlap != 0 {
                self.overlap = series.overlap;
            }

            // Copy a series gap to the parent chart.
            if series.gap != 150 {
                self.gap = series.gap;
            }

            // Write the c:idx element.
            self.write_idx(self.base_series_index + index);

            // Write the c:order element.
            self.write_order(self.base_series_index + index);

            self.write_series_title(&series.title);

            // Write the c:spPr element.
            self.write_sp_pr(&series.format);

            if let Some(marker) = &series.marker {
                if !marker.automatic {
                    // Write the c:marker element.
                    self.write_marker(marker);
                }
            }

            // Write the c:invertIfNegative element.
            if series.invert_if_negative {
                self.write_invert_if_negative();
            }

            // Write the point formatting for the series.
            if !series.points.is_empty() {
                self.write_d_pt(&series.points, max_points);
            }

            if let Some(data_label) = &series.data_label {
                // Write the c:dLbls element.
                self.write_data_labels(data_label, &series.custom_data_labels, max_points);
            }

            if series.trendline.trend_type != ChartTrendlineType::None {
                // Write the c:trendline element.
                self.write_trendline(&series.trendline);
            }

            if self.chart_group_type == ChartType::Bar {
                if let Some(error_bars) = &series.x_error_bars {
                    // Write the c:errBars element.
                    self.write_error_bar("", error_bars);
                }
            } else if self.chart_group_type == ChartType::Column {
                if let Some(error_bars) = &series.y_error_bars {
                    // Write the c:errBars element.
                    self.write_error_bar("", error_bars);
                }
            } else if let Some(error_bars) = &series.y_error_bars {
                // Write the c:errBars element.
                self.write_error_bar("y", error_bars);
            }

            // Write the c:cat element.
            if series.category_range.has_data() {
                // We only set a default num format for non-string categories.
                self.category_has_num_format =
                    series.category_range.cache.cache_type != ChartRangeCacheDataType::String;
                self.write_cat(&series.category_range);
            }

            // Write the c:val element.
            self.write_val(&series.value_range);

            if !series.inverted_color.is_auto_or_default() {
                // Write the c:extLst element for the inverted fill color.
                self.write_extension_list(series.inverted_color);
            }

            // Write the c:smooth element.
            if self.chart_group_type == ChartType::Line {
                if let Some(smooth) = series.smooth {
                    if smooth {
                        self.write_smooth();
                    }
                }
            }

            self.writer.xml_end_tag("c:ser");
        }
    }

    // Write the <c:ser> element for scatter charts.
    fn write_scatter_series(&mut self) {
        for (index, series) in self.series.clone().iter_mut().enumerate() {
            let max_points = series.value_range.number_of_points();

            self.writer.xml_start_tag_only("c:ser");

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
                self.write_d_pt(&series.points, max_points);
            }

            // Write the c:dLbls element.
            if let Some(data_label) = &series.data_label {
                self.write_data_labels(data_label, &series.custom_data_labels, max_points);
            }

            // Write the c:trendline element.
            if series.trendline.trend_type != ChartTrendlineType::None {
                self.write_trendline(&series.trendline);
            }

            // Write the X-Axis c:errBars element.
            if let Some(error_bars) = &series.x_error_bars {
                self.write_error_bar("x", error_bars);
            }

            // Write the Y-Axis the c:errBars element.
            if let Some(error_bars) = &series.y_error_bars {
                self.write_error_bar("y", error_bars);
            }

            self.write_x_val(&series.category_range);

            self.write_y_val(&series.value_range);

            // Write the c:smooth element.
            if self.chart_group_type == ChartType::Scatter {
                if let Some(smooth) = series.smooth {
                    if smooth {
                        self.write_smooth();
                    }
                } else if self.chart_type == ChartType::ScatterSmooth
                    || self.chart_type == ChartType::ScatterSmoothWithMarkers
                {
                    // The ScatterSmooth charts have a default smooth element if
                    // one hasn't been set/unset by the user.
                    self.write_smooth();
                }
            }

            self.writer.xml_end_tag("c:ser");
        }
    }

    // Write the <c:dPt> element.
    fn write_d_pt(&mut self, points: &[ChartPoint], max_points: usize) {
        let has_marker =
            self.chart_group_type == ChartType::Scatter || self.chart_group_type == ChartType::Line;

        // Write the point formatting for the series.
        for (index, point) in points.iter().enumerate() {
            if index >= max_points {
                break;
            }

            if point.is_not_default() {
                self.writer.xml_start_tag_only("c:dPt");
                self.write_idx(index);

                if has_marker {
                    self.writer.xml_start_tag_only("c:marker");
                }

                // Write the c:spPr formatting element.
                self.write_sp_pr(&point.format);

                if has_marker {
                    self.writer.xml_end_tag("c:marker");
                }

                self.writer.xml_end_tag("c:dPt");
            }
        }
    }

    // Write the <c:idx> element.
    fn write_idx(&mut self, index: usize) {
        let attributes = [("val", index.to_string())];

        self.writer.xml_empty_tag("c:idx", &attributes);
    }

    // Write the <c:order> element.
    fn write_order(&mut self, index: usize) {
        let attributes = [("val", index.to_string())];

        self.writer.xml_empty_tag("c:order", &attributes);
    }

    // Write the <c:invertIfNegative> element.
    fn write_invert_if_negative(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:invertIfNegative", &attributes);
    }

    // Write the <c:extLst> element for inverted fill colors.
    fn write_extension_list(&mut self, color: Color) {
        let attributes1 = [
            ("uri", "{6F2FDCE9-48DA-4B69-8628-5D25D57E5C99}"),
            (
                "xmlns:c14",
                "http://schemas.microsoft.com/office/drawing/2007/8/2/chart",
            ),
        ];
        let attributes2 = [(
            "xmlns:c14",
            "http://schemas.microsoft.com/office/drawing/2007/8/2/chart",
        )];

        self.writer.xml_start_tag_only("c:extLst");
        self.writer.xml_start_tag("c:ext", &attributes1);
        self.writer.xml_start_tag_only("c14:invertSolidFillFmt");
        self.writer.xml_start_tag("c14:spPr", &attributes2);

        self.write_a_solid_fill(color, 0);

        self.writer.xml_end_tag("c14:spPr");
        self.writer.xml_end_tag("c14:invertSolidFillFmt");
        self.writer.xml_end_tag("c:ext");
        self.writer.xml_end_tag("c:extLst");
    }

    // Write the <c:cat> element.
    fn write_cat(&mut self, range: &ChartRange) {
        self.writer.xml_start_tag_only("c:cat");

        self.write_cache_ref(range, false);

        self.writer.xml_end_tag("c:cat");
    }

    // Write the <c:val> element.
    fn write_val(&mut self, range: &ChartRange) {
        self.writer.xml_start_tag_only("c:val");

        self.write_cache_ref(range, true);

        self.writer.xml_end_tag("c:val");
    }

    // Write the <c:xVal> element for scatter charts.
    fn write_x_val(&mut self, range: &ChartRange) {
        self.writer.xml_start_tag_only("c:xVal");

        self.write_cache_ref(range, false);

        self.writer.xml_end_tag("c:xVal");
    }

    // Write the <c:yVal> element for scatter charts.
    fn write_y_val(&mut self, range: &ChartRange) {
        self.writer.xml_start_tag_only("c:yVal");

        self.write_cache_ref(range, true);

        self.writer.xml_end_tag("c:yVal");
    }

    // Write the <c:numRef> or <c:strRef> elements. Value range must be written
    // as a numRef where strings are treated as zero.
    fn write_cache_ref(&mut self, range: &ChartRange, is_num_only: bool) {
        if range.cache.cache_type == ChartRangeCacheDataType::String && !is_num_only {
            self.write_str_ref(range);
        } else {
            self.write_num_ref(range);
        }
    }

    // Write the <c:numRef> element.
    fn write_num_ref(&mut self, range: &ChartRange) {
        self.writer.xml_start_tag_only("c:numRef");

        // Write the c:f element.
        self.write_range_formula(&range.formula_abs());

        // Write the c:numCache element.
        if range.cache.has_data() {
            self.write_num_cache(&range.cache);
        }

        self.writer.xml_end_tag("c:numRef");
    }

    // Write the <c:strRef> element.
    fn write_str_ref(&mut self, range: &ChartRange) {
        self.writer.xml_start_tag_only("c:strRef");

        // Write the c:f element.
        self.write_range_formula(&range.formula_abs());

        // Write the c:strCache element.
        if range.cache.has_data() {
            self.write_str_cache(&range.cache);
        }

        self.writer.xml_end_tag("c:strRef");
    }

    // Write the <c:numCache> element.
    fn write_num_cache(&mut self, cache: &ChartRangeCacheData) {
        self.writer.xml_start_tag_only("c:numCache");

        // Write the c:formatCode element.
        if cache.cache_type == ChartRangeCacheDataType::Date {
            self.write_format_code("dd/mm/yyyy");
        } else {
            self.write_format_code("General");
        }

        // Write the c:ptCount element.
        self.write_pt_count(cache.data.len());

        // Write the c:pt elements.
        for (index, value) in cache.data.iter().enumerate() {
            if !value.is_empty() {
                // Non numeric values in value/number caches are treated as zero
                // by Excel.
                if value.parse::<f64>().is_err() {
                    self.write_pt(index, "0");
                } else {
                    self.write_pt(index, value);
                }
            }
        }

        self.writer.xml_end_tag("c:numCache");
    }

    // Write the <c:strCache> element.
    fn write_str_cache(&mut self, cache: &ChartRangeCacheData) {
        self.writer.xml_start_tag_only("c:strCache");

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
        self.writer.xml_data_element_only("c:f", formula);
    }

    // Write the <c:formatCode> element.
    fn write_format_code(&mut self, format_code: &str) {
        self.writer
            .xml_data_element_only("c:formatCode", format_code);
    }

    // Write the <c:ptCount> element.
    fn write_pt_count(&mut self, count: usize) {
        let attributes = [("val", count.to_string())];

        self.writer.xml_empty_tag("c:ptCount", &attributes);
    }

    // Write the <c:pt> element.
    fn write_pt(&mut self, index: usize, value: &str) {
        let attributes = [("idx", index.to_string())];

        self.writer.xml_start_tag("c:pt", &attributes);
        self.writer.xml_data_element_only("c:v", value);
        self.writer.xml_end_tag("c:pt");
    }

    // Write both <c:axId> elements.
    fn write_ax_ids(&mut self) {
        self.write_ax_id(self.axis_ids.0);
        self.write_ax_id(self.axis_ids.1);
    }

    // Write the <c:axId> element.
    fn write_ax_id(&mut self, axis_id: u32) {
        let attributes = [("val", axis_id.to_string())];

        self.writer.xml_empty_tag("c:axId", &attributes);
    }

    // -----------------------------------------------------------------------
    // Category Axis.
    // -----------------------------------------------------------------------

    // Write the <c:catAx> element.
    fn write_cat_ax(&mut self) {
        self.writer.xml_start_tag_only("c:catAx");

        self.write_ax_id(self.axis_ids.0);

        // Write the c:scaling element.
        self.write_scaling(&self.x_axis.clone());

        if self.x_axis.is_hidden {
            self.write_delete();
        }

        // Write the c:axPos element.
        self.write_ax_pos(
            self.x_axis.axis_position,
            self.y_axis.reverse,
            self.y_axis.crossing,
        );

        self.write_major_gridlines(self.x_axis.clone());
        self.write_minor_gridlines(self.x_axis.clone());

        // Write the c:title element.
        self.write_chart_title(&self.x_axis.title.clone());

        // Write the c:numFmt element.
        if !self.x_axis.num_format.is_empty() {
            self.write_number_format(
                &self.x_axis.num_format.clone(),
                self.x_axis.num_format_linked_to_source,
            );
        } else if self.category_has_num_format {
            self.write_number_format("General", true);
        }

        // Write the c:majorTickMark element.
        if let Some(tick_type) = self.x_axis.major_tick_type {
            self.write_major_tick_mark(tick_type);
        }

        // Write the c:minorTickMark element.
        if let Some(tick_type) = self.x_axis.minor_tick_type {
            self.write_minor_tick_mark(tick_type);
        }

        // Write the c:tickLblPos element.
        self.write_tick_label_position(self.x_axis.label_position);

        if self.x_axis.format.has_formatting() {
            // Write the c:spPr formatting element.
            self.write_sp_pr(&self.x_axis.format.clone());
        }

        // Write the axis font elements.
        if let Some(font) = &self.x_axis.font {
            self.write_axis_font(&font.clone());
        }

        // Write the c:crossAx element.
        self.write_cross_ax(self.axis_ids.1);

        // Write the c:crosses element. Note, the X crossing comes from the Y
        // axis.
        match self.y_axis.crossing {
            ChartAxisCrossing::Automatic | ChartAxisCrossing::Min | ChartAxisCrossing::Max => {
                self.write_crosses(&self.y_axis.crossing.to_string());
            }
            ChartAxisCrossing::AxisValue(_) => {
                self.write_crosses_at(&self.y_axis.crossing.to_string());
            }
            ChartAxisCrossing::CategoryNumber(_) => {
                // Ignore Category crossing on a Value axis.
                self.write_crosses(&ChartAxisCrossing::Automatic.to_string());
            }
        }

        // Write the c:auto element.
        if !self.x_axis.automatic {
            self.write_auto();
        }

        // Write the c:lblAlgn element.
        self.write_lbl_algn(&self.x_axis.label_alignment.to_string());

        // Write the c:lblOffset element.
        self.write_lbl_offset();

        // Write the c:tickLblSkip element.
        if self.x_axis.label_interval > 1 {
            self.write_tick_lbl_skip(self.x_axis.label_interval);
        }

        // Write the c:tickMarkSkip element.
        if self.x_axis.tick_interval > 1 {
            self.write_tick_mark_skip(self.x_axis.tick_interval);
        }

        self.writer.xml_end_tag("c:catAx");
    }

    // -----------------------------------------------------------------------
    // Date Axis.
    // -----------------------------------------------------------------------

    // Write the <c:dateAx> element.
    fn write_date_ax(&mut self) {
        self.writer.xml_start_tag_only("c:dateAx");

        self.write_ax_id(self.axis_ids.0);

        // Write the c:scaling element.
        self.write_scaling(&self.x_axis.clone());

        if self.x_axis.is_hidden {
            self.write_delete();
        }

        // Write the c:axPos element.
        self.write_ax_pos(
            self.x_axis.axis_position,
            self.y_axis.reverse,
            self.y_axis.crossing,
        );

        self.write_major_gridlines(self.x_axis.clone());
        self.write_minor_gridlines(self.x_axis.clone());

        // Write the c:title element.
        self.write_chart_title(&self.x_axis.title.clone());

        // Write the c:numFmt element.
        if !self.x_axis.num_format.is_empty() {
            self.write_number_format(
                &self.x_axis.num_format.clone(),
                self.x_axis.num_format_linked_to_source,
            );
        } else if self.category_has_num_format {
            self.write_number_format("dd/mm/yyyy", true);
        }

        // Write the c:majorTickMark element.
        if let Some(tick_type) = self.x_axis.major_tick_type {
            self.write_major_tick_mark(tick_type);
        }

        // Write the c:minorTickMark element.
        if let Some(tick_type) = self.x_axis.minor_tick_type {
            self.write_minor_tick_mark(tick_type);
        }

        // Write the c:tickLblPos element.
        self.write_tick_label_position(self.x_axis.label_position);

        if self.x_axis.format.has_formatting() {
            // Write the c:spPr formatting element.
            self.write_sp_pr(&self.x_axis.format.clone());
        }

        // Write the axis font elements.
        if let Some(font) = &self.x_axis.font {
            self.write_axis_font(&font.clone());
        }

        // Write the c:crossAx element.
        self.write_cross_ax(self.axis_ids.1);

        // Write the c:crosses element. Note, the X crossing comes from the Y
        // axis.
        match self.y_axis.crossing {
            ChartAxisCrossing::Automatic | ChartAxisCrossing::Min | ChartAxisCrossing::Max => {
                self.write_crosses(&self.y_axis.crossing.to_string());
            }
            ChartAxisCrossing::AxisValue(_) => {
                self.write_crosses_at(&self.y_axis.crossing.to_string());
            }
            ChartAxisCrossing::CategoryNumber(_) => {
                // Ignore Category crossing on a Value axis.
                self.write_crosses(&ChartAxisCrossing::Automatic.to_string());
            }
        }

        // Write the c:auto element.
        if self.x_axis.automatic {
            self.write_auto();
        }

        // Write the c:lblOffset element.
        self.write_lbl_offset();

        // Write the c:tickLblSkip element.
        if self.x_axis.label_interval > 1 {
            self.write_tick_lbl_skip(self.x_axis.label_interval);
        }

        // Write the c:tickMarkSkip element.
        if self.x_axis.tick_interval > 1 {
            self.write_tick_mark_skip(self.x_axis.tick_interval);
        }

        // Write the c:majorUnit element.
        if !self.x_axis.major_unit.is_empty() {
            self.write_major_unit(self.x_axis.major_unit.clone());
        }

        // Write the c:majorTimeUnit element.
        if let Some(unit) = self.x_axis.major_unit_date_type {
            self.write_major_time_unit(unit);
        }

        // Write the c:minorUnit element.
        if !self.x_axis.minor_unit.is_empty() {
            self.write_minor_unit(self.x_axis.minor_unit.clone());
        }

        // Write the c:minorTimeUnit element.
        if let Some(unit) = self.x_axis.minor_unit_date_type {
            self.write_minor_time_unit(unit);
        }

        self.writer.xml_end_tag("c:dateAx");
    }

    // -----------------------------------------------------------------------
    // Value Axis.
    // -----------------------------------------------------------------------

    // Write the <c:valAx> element.
    fn write_val_ax(&mut self) {
        self.writer.xml_start_tag_only("c:valAx");

        self.write_ax_id(self.axis_ids.1);

        // Write the c:scaling element.
        self.write_scaling(&self.y_axis.clone());

        if self.y_axis.is_hidden {
            self.write_delete();
        }
        // Write the c:axPos element.
        self.write_ax_pos(
            self.y_axis.axis_position,
            self.x_axis.reverse,
            self.x_axis.crossing,
        );

        // Write the Gridlines elements.
        self.write_major_gridlines(self.y_axis.clone());
        self.write_minor_gridlines(self.y_axis.clone());

        // Write the c:title element.
        self.write_chart_title(&self.y_axis.title.clone());

        // Write the c:numFmt element.
        if self.y_axis.num_format.is_empty() {
            self.write_number_format(&self.default_num_format.clone(), true);
        } else {
            self.write_number_format(
                &self.y_axis.num_format.clone(),
                self.y_axis.num_format_linked_to_source,
            );
        }

        // Write the c:majorTickMark element.
        if let Some(position) = self.y_axis.major_tick_type {
            self.write_major_tick_mark(position);
        }

        // Write the c:minorTickMark element.
        if let Some(position) = self.y_axis.minor_tick_type {
            self.write_minor_tick_mark(position);
        }

        // Write the c:tickLblPos element.
        self.write_tick_label_position(self.y_axis.label_position);

        if self.y_axis.format.has_formatting() {
            // Write the c:spPr formatting element.
            self.write_sp_pr(&self.y_axis.format.clone());
        }

        // Write the axis font elements.
        if let Some(font) = &self.y_axis.font {
            self.write_axis_font(&font.clone());
        }

        // Write the c:crossAx element.
        self.write_cross_ax(self.axis_ids.0);

        // Write the c:crosses element. Note, the Y crossing comes from the X
        // axis.
        match self.x_axis.crossing {
            ChartAxisCrossing::Automatic | ChartAxisCrossing::Min | ChartAxisCrossing::Max => {
                self.write_crosses(&self.x_axis.crossing.to_string());
            }
            ChartAxisCrossing::CategoryNumber(_) | ChartAxisCrossing::AxisValue(_) => {
                self.write_crosses_at(&self.x_axis.crossing.to_string());
            }
        }

        // Write the c:crossBetween element.
        self.write_cross_between(self.x_axis.position_between_ticks);

        // Write the c:majorUnit element.
        if self.y_axis.axis_type != ChartAxisType::Category && !self.y_axis.major_unit.is_empty() {
            self.write_major_unit(self.y_axis.major_unit.clone());
        }

        // Write the c:minorUnit element.
        if self.y_axis.axis_type != ChartAxisType::Category && !self.y_axis.minor_unit.is_empty() {
            self.write_minor_unit(self.y_axis.minor_unit.clone());
        }

        // Write the c:dispUnits element.
        if self.y_axis.display_units_type != ChartAxisDisplayUnitType::None {
            self.write_disp_units(
                self.y_axis.display_units_type,
                self.y_axis.display_units_visible,
            );
        }

        self.writer.xml_end_tag("c:valAx");
    }

    // -----------------------------------------------------------------------
    // Category Value Axis. Only for Scatter charts.
    // -----------------------------------------------------------------------

    // Write the category <c:valAx> element for scatter charts.
    fn write_cat_val_ax(&mut self) {
        self.writer.xml_start_tag_only("c:valAx");

        self.write_ax_id(self.axis_ids.0);

        // Write the c:scaling element.
        self.write_scaling(&self.x_axis.clone());

        if self.x_axis.is_hidden {
            self.write_delete();
        }

        // Write the c:axPos element.
        self.write_ax_pos(
            self.x_axis.axis_position,
            self.y_axis.reverse,
            self.y_axis.crossing,
        );

        // Write the Gridlines elements.
        self.write_major_gridlines(self.x_axis.clone());
        self.write_minor_gridlines(self.x_axis.clone());

        // Write the c:title element.
        self.write_chart_title(&self.x_axis.title.clone());

        // Write the c:numFmt element.
        if self.x_axis.num_format.is_empty() {
            self.write_number_format(&self.default_num_format.clone(), true);
        } else {
            self.write_number_format(
                &self.x_axis.num_format.clone(),
                self.x_axis.num_format_linked_to_source,
            );
        }

        // Write the c:majorTickMark element.
        if let Some(position) = self.x_axis.major_tick_type {
            self.write_major_tick_mark(position);
        }

        // Write the c:minorTickMark element.
        if let Some(position) = self.x_axis.minor_tick_type {
            self.write_minor_tick_mark(position);
        }

        // Write the c:tickLblPos element.
        self.write_tick_label_position(self.x_axis.label_position);

        if self.x_axis.format.has_formatting() {
            // Write the c:spPr formatting element.
            self.write_sp_pr(&self.x_axis.format.clone());
        }

        // Write the axis font elements.
        if let Some(font) = &self.x_axis.font {
            self.write_axis_font(&font.clone());
        }

        // Write the c:crossAx element.
        self.write_cross_ax(self.axis_ids.1);

        // Write the c:crosses element. Note, the X crossing comes from the Y
        // axis.
        match self.y_axis.crossing {
            ChartAxisCrossing::Automatic | ChartAxisCrossing::Min | ChartAxisCrossing::Max => {
                self.write_crosses(&self.y_axis.crossing.to_string());
            }
            ChartAxisCrossing::CategoryNumber(_) | ChartAxisCrossing::AxisValue(_) => {
                self.write_crosses_at(&self.y_axis.crossing.to_string());
            }
        }

        // Write the c:crossBetween element.
        self.write_cross_between(self.y_axis.position_between_ticks);

        // Write the c:majorUnit element.
        if self.x_axis.axis_type != ChartAxisType::Category && !self.x_axis.major_unit.is_empty() {
            self.write_major_unit(self.x_axis.major_unit.clone());
        }

        // Write the c:minorUnit element.
        if self.x_axis.axis_type != ChartAxisType::Category && !self.x_axis.minor_unit.is_empty() {
            self.write_minor_unit(self.x_axis.minor_unit.clone());
        }

        // Write the c:dispUnits element.
        if self.x_axis.display_units_type != ChartAxisDisplayUnitType::None {
            self.write_disp_units(
                self.x_axis.display_units_type,
                self.x_axis.display_units_visible,
            );
        }

        self.writer.xml_end_tag("c:valAx");
    }

    // Write the <c:scaling> element.
    fn write_scaling(&mut self, axis: &ChartAxis) {
        self.writer.xml_start_tag_only("c:scaling");

        // Write the c:logBase element.
        if axis.axis_type != ChartAxisType::Category && axis.log_base >= 2 {
            self.write_log_base(axis.log_base);
        }

        // Write the c:orientation element.
        self.write_orientation(axis.reverse);

        // Write the c:max element.
        if axis.axis_type != ChartAxisType::Category && !axis.max.is_empty() {
            self.write_max(&axis.max);
        }

        // Write the c:min element.
        if axis.axis_type != ChartAxisType::Category && !axis.min.is_empty() {
            self.write_min(&axis.min);
        }

        self.writer.xml_end_tag("c:scaling");
    }

    // Write the <c:logBase> element.
    fn write_log_base(&mut self, base: u16) {
        let attributes = [("val", base.to_string())];

        self.writer.xml_empty_tag("c:logBase", &attributes);
    }

    // Write the <c:orientation> element.
    fn write_orientation(&mut self, reverse: bool) {
        let attributes = if reverse {
            [("val", "maxMin")]
        } else {
            [("val", "minMax")]
        };

        self.writer.xml_empty_tag("c:orientation", &attributes);
    }

    // Write the <c:max> element.
    fn write_max(&mut self, max: &str) {
        let attributes = [("val", max.to_string())];

        self.writer.xml_empty_tag("c:max", &attributes);
    }

    // Write the <c:min> element.
    fn write_min(&mut self, min: &str) {
        let attributes = [("val", min.to_string())];

        self.writer.xml_empty_tag("c:min", &attributes);
    }

    // Write the <c:axPos> element.
    fn write_ax_pos(
        &mut self,
        position: ChartAxisPosition,
        reverse: bool,
        crossing: ChartAxisCrossing,
    ) {
        let mut position = position;

        if reverse || crossing == ChartAxisCrossing::Max {
            position = position.reverse();
        }

        let attributes = [("val", position.to_string())];

        self.writer.xml_empty_tag("c:axPos", &attributes);
    }

    // Write the <c:numFmt> element.
    fn write_number_format(&mut self, format: &str, linked: bool) {
        let attributes = [
            ("formatCode", format.to_string()),
            ("sourceLinked", linked.to_xml_bool()),
        ];

        self.writer.xml_empty_tag("c:numFmt", &attributes);
    }

    // Write the <c:majorGridlines> element.
    fn write_major_gridlines(&mut self, axis: ChartAxis) {
        if axis.major_gridlines {
            if let Some(line) = &axis.major_gridlines_line {
                self.writer.xml_start_tag_only("c:majorGridlines");
                self.writer.xml_start_tag_only("c:spPr");

                // Write the a:ln element.
                self.write_a_ln(line);

                self.writer.xml_end_tag("c:spPr");
                self.writer.xml_end_tag("c:majorGridlines");
            } else {
                self.writer.xml_empty_tag_only("c:majorGridlines");
            }
        }
    }

    // Write the <c:minorGridlines> element.
    fn write_minor_gridlines(&mut self, axis: ChartAxis) {
        if axis.minor_gridlines {
            if let Some(line) = &axis.minor_gridlines_line {
                self.writer.xml_start_tag_only("c:minorGridlines");
                self.writer.xml_start_tag_only("c:spPr");

                // Write the a:ln element.
                self.write_a_ln(line);

                self.writer.xml_end_tag("c:spPr");
                self.writer.xml_end_tag("c:minorGridlines");
            } else {
                self.writer.xml_empty_tag_only("c:minorGridlines");
            }
        }
    }

    // Write the <c:tickLblPos> element.
    fn write_tick_label_position(&mut self, position: ChartAxisLabelPosition) {
        let attributes = [("val", position.to_string())];

        self.writer.xml_empty_tag("c:tickLblPos", &attributes);
    }

    // Write the <c:crossAx> element.
    fn write_cross_ax(&mut self, axis_id: u32) {
        let attributes = [("val", axis_id.to_string())];

        self.writer.xml_empty_tag("c:crossAx", &attributes);
    }

    // Write the <c:crosses> element.
    fn write_crosses(&mut self, crossing: &str) {
        let attributes = [("val", crossing)];

        self.writer.xml_empty_tag("c:crosses", &attributes);
    }

    // Write the <c:crossesAt> element.
    fn write_crosses_at(&mut self, crossing: &str) {
        let attributes = [("val", crossing)];

        self.writer.xml_empty_tag("c:crossesAt", &attributes);
    }

    // Write the <c:auto> element.
    fn write_auto(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:auto", &attributes);
    }

    // Write the <c:lblAlgn> element.
    fn write_lbl_algn(&mut self, position: &str) {
        let attributes = [("val", position)];

        self.writer.xml_empty_tag("c:lblAlgn", &attributes);
    }

    // Write the <c:lblOffset> element.
    fn write_lbl_offset(&mut self) {
        let attributes = [("val", "100")];

        self.writer.xml_empty_tag("c:lblOffset", &attributes);
    }

    // Write the <c:crossBetween> element.
    fn write_cross_between(&mut self, position_between_ticks: bool) {
        let attributes = if position_between_ticks {
            [("val", "between")]
        } else {
            [("val", "midCat")]
        };

        self.writer.xml_empty_tag("c:crossBetween", &attributes);
    }

    // Write the <c:tickLblSkip> element.
    fn write_tick_lbl_skip(&mut self, units: u16) {
        let attributes = [("val", units.to_string())];

        self.writer.xml_empty_tag("c:tickLblSkip", &attributes);
    }

    // Write the <c:tickMarkSkip> element.
    fn write_tick_mark_skip(&mut self, units: u16) {
        let attributes = [("val", units.to_string())];

        self.writer.xml_empty_tag("c:tickMarkSkip", &attributes);
    }

    // Write the <c:majorUnit> element.
    fn write_major_unit(&mut self, value: String) {
        let attributes = [("val", value)];

        self.writer.xml_empty_tag("c:majorUnit", &attributes);
    }

    // Write the <c:minorUnit> element.
    fn write_minor_unit(&mut self, value: String) {
        let attributes = [("val", value)];

        self.writer.xml_empty_tag("c:minorUnit", &attributes);
    }

    // Write the <c:majorTimeUnit> element.
    fn write_major_time_unit(&mut self, units: ChartAxisDateUnitType) {
        let attributes = [("val", units.to_string())];

        self.writer.xml_empty_tag("c:majorTimeUnit", &attributes);
    }

    // Write the <c:minorTimeUnit> element.
    fn write_minor_time_unit(&mut self, units: ChartAxisDateUnitType) {
        let attributes = [("val", units.to_string())];

        self.writer.xml_empty_tag("c:minorTimeUnit", &attributes);
    }

    // Write the <c:dispUnits> element.
    fn write_disp_units(&mut self, units: ChartAxisDisplayUnitType, visible: bool) {
        self.writer.xml_start_tag_only("c:dispUnits");

        // Write the c:builtInUnit element.
        self.write_built_in_unit(units);

        // Write the c:dispUnitsLbl element.
        if visible {
            self.write_disp_units_lbl();
        }

        self.writer.xml_end_tag("c:dispUnits");
    }

    // Write the <c:builtInUnit> element.
    fn write_built_in_unit(&mut self, units: ChartAxisDisplayUnitType) {
        let attributes = [("val", units.to_string())];

        self.writer.xml_empty_tag("c:builtInUnit", &attributes);
    }

    // Write the <c:dispUnitsLbl> element.
    fn write_disp_units_lbl(&mut self) {
        self.writer.xml_start_tag_only("c:dispUnitsLbl");

        // Write the c:layout element.
        self.write_layout();

        self.writer.xml_end_tag("c:dispUnitsLbl");
    }

    // Write the <c:legend> element.
    fn write_legend(&mut self) {
        if self.legend.hidden {
            return;
        }

        self.writer.xml_start_tag_only("c:legend");

        // Write the c:legendPos element.
        self.write_legend_pos();

        // Check for series and trendlines that should be deleted/hidden from
        // the legend.
        let deleted_entries = self.deleted_legend_entries();

        if !deleted_entries.is_empty() {
            for index in deleted_entries {
                // Write the c:legendEntry element.
                self.write_legend_entry(index);
            }
        }

        // Write the c:layout element.
        self.write_layout();

        // Write the c:spPr formatting element.
        self.write_sp_pr(&self.legend.format.clone());

        // Write the c:overlay element.
        self.write_overlay();

        // Pie/Doughnut charts set the "rtl" flag to "0" in the legend font even
        // though "0" is implied. To match Excel output we set it if it hasn't
        // been set by the user.
        if self.chart_type == ChartType::Pie || self.chart_type == ChartType::Doughnut {
            match &mut self.legend.font {
                Some(font) => {
                    if font.right_to_left.is_none() {
                        font.set_right_to_left(false);
                    }
                }
                None => {
                    let mut font = ChartFont::new();
                    font.set_right_to_left(false);
                    self.legend.font = Some(font);
                }
            };
        }

        if let Some(font) = &self.legend.font {
            // Write the c:txPr element.
            self.write_tx_pr(&font.clone(), false);
        }

        self.writer.xml_end_tag("c:legend");
    }

    // Write the <c:legendPos> element.
    fn write_legend_pos(&mut self) {
        let attributes = [("val", self.legend.position.to_string())];

        self.writer.xml_empty_tag("c:legendPos", &attributes);
    }

    // Write the <c:legendEntry> element.
    fn write_legend_entry(&mut self, index: usize) {
        self.writer.xml_start_tag_only("c:legendEntry");

        // Write the c:idx element.
        self.write_idx(index);

        // Write the c:delete element.
        self.write_delete();

        self.writer.xml_end_tag("c:legendEntry");
    }

    // Write the <c:overlay> element.
    fn write_overlay(&mut self) {
        if !self.legend.has_overlay {
            return;
        }

        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:overlay", &attributes);
    }

    // Write the <c:plotVisOnly> element.
    fn write_plot_vis_only(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:plotVisOnly", &attributes);
    }

    // Write the <c:printSettings> element.
    fn write_print_settings(&mut self) {
        self.writer.xml_start_tag_only("c:printSettings");

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
        self.writer.xml_empty_tag_only("c:headerFooter");
    }

    // Write the <c:pageMargins> element.
    fn write_page_margins(&mut self) {
        let attributes = [
            ("b", "0.75"),
            ("l", "0.7"),
            ("r", "0.7"),
            ("t", "0.75"),
            ("header", "0.3"),
            ("footer", "0.3"),
        ];

        self.writer.xml_empty_tag("c:pageMargins", &attributes);
    }

    // Write the <c:pageSetup> element.
    fn write_page_setup(&mut self) {
        self.writer.xml_empty_tag_only("c:pageSetup");
    }

    // Write the <c:marker> element.
    fn write_marker_value(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:marker", &attributes);
    }

    // Write the <c:marker> element.
    fn write_marker(&mut self, marker: &ChartMarker) {
        self.writer.xml_start_tag_only("c:marker");

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

    // Write the <c:dLbls> element.
    fn write_data_labels(
        &mut self,
        data_label: &ChartDataLabel,
        custom_data_labels: &[ChartDataLabel],
        max_points: usize,
    ) {
        self.writer.xml_start_tag_only("c:dLbls");

        if !custom_data_labels.is_empty() {
            self.write_custom_data_labels(custom_data_labels, max_points);
        }

        // Write the main elements of a data label.
        self.write_data_label(data_label);

        self.writer.xml_end_tag("c:dLbls");
    }

    // Write the <c:dLbl> element.
    fn write_custom_data_labels(&mut self, data_labels: &[ChartDataLabel], max_points: usize) {
        // Write the point formatting for the series.
        for (index, data_label) in data_labels.iter().enumerate() {
            let mut write_layout = true;

            if index >= max_points {
                break;
            }

            if data_label.is_default() {
                continue;
            }

            self.writer.xml_start_tag_only("c:dLbl");
            self.write_idx(index);

            if data_label.is_hidden {
                // Write the c:delete element.
                self.write_delete();
            } else {
                // Add empty "c:spPr", as required, for Excel compatibility.
                if !data_label.format.has_formatting() {
                    if let Some(font) = &data_label.font {
                        if font.color.is_auto_or_default() {
                            self.writer.xml_empty_tag_only("c:spPr");
                        }
                    }
                }

                // If a custom point has a font then it may need to be applied
                // to the title and/or the label.
                let mut data_label = data_label.clone();
                data_label.is_custom = true;

                if let Some(font) = &mut data_label.font {
                    font.has_baseline = false;
                    write_layout = false;
                }

                if !data_label.title.name.is_empty() || data_label.title.range.has_data() {
                    if let Some(font) = &data_label.font {
                        data_label.title.set_font(font);
                        data_label.title.font.has_baseline = false;

                        if !data_label.title.name.is_empty() {
                            data_label.font = None;
                        }

                        write_layout = true;
                    }
                }

                if write_layout {
                    // Write the c:layout element.
                    self.write_layout();
                }

                // Write the c:tx element.
                if !data_label.title.name.is_empty() {
                    self.write_tx_rich(&data_label.title);
                } else if data_label.title.range.has_data() {
                    self.write_tx_formula(&data_label.title);
                }

                // Write the main elements of a data label.
                self.write_data_label(&data_label);
            }

            self.writer.xml_end_tag("c:dLbl");
        }
    }

    fn write_data_label(&mut self, data_label: &ChartDataLabel) {
        if !data_label.num_format.is_empty() {
            // Write the c:numFmt element.
            self.write_number_format(&data_label.num_format, false);
        }

        // Write the c:spPr formatting element.
        self.write_sp_pr(&data_label.format);

        if let Some(font) = &data_label.font {
            // Write the c:txPr element.
            self.write_tx_pr(&font.clone(), false);
        }

        if data_label.position != ChartDataLabelPosition::Default
            && data_label.position != self.default_label_position
        {
            // Write the c:dLblPos element.
            self.write_d_lbl_pos(data_label.position);
        }

        if data_label.show_legend_key {
            // Write the c:showLegendKey element.
            self.write_show_legend_key();
        }

        // Ensure at least one display option is set.
        if data_label.show_value
            || (!data_label.is_custom
                && !data_label.show_category_name
                && !data_label.show_percentage)
        {
            // Write the c:showVal element.
            self.write_show_val();
        }

        if data_label.show_category_name {
            // Write the c:showCatName element.
            self.write_show_category_name();
        }

        if data_label.show_series_name {
            // Write the c:showSerName element.
            self.write_show_series_name();
        }

        if data_label.show_percentage {
            // Write the c:showPercent element.
            self.write_show_percent();
        }

        if data_label.separator != ',' {
            // Write the c:separator element.
            self.write_separator(data_label.separator);
        }

        if data_label.show_leader_lines {
            match self.chart_group_type {
                // Write the c:showLeaderLines element.
                ChartType::Pie | ChartType::Doughnut => {
                    self.write_show_leader_lines_2007();
                }
                _ => {
                    self.write_show_leader_lines_2015();
                }
            }
        }
    }

    // Write the <c:trendline> element.
    fn write_trendline(&mut self, trendline: &ChartTrendline) {
        self.writer.xml_start_tag_only("c:trendline");

        if !trendline.name.is_empty() {
            // Write the c:name element.
            self.write_trendline_name(&trendline.name);
        }

        // Write the c:spPr formatting element.
        self.write_sp_pr(&trendline.format);

        // Write the c:trendlineType element.
        self.write_trendline_type(trendline);

        if let ChartTrendlineType::Polynomial(order) = trendline.trend_type {
            self.write_order(order as usize);
        }

        if let ChartTrendlineType::MovingAverage(period) = trendline.trend_type {
            // Write the c:period element.
            self.write_trendline_period(period);
        }

        if trendline.forward_period > 0.0 {
            // Write the c:forward element.
            self.write_trendline_forward(trendline.forward_period);
        }

        if trendline.backward_period > 0.0 {
            // Write the c:backward element.
            self.write_trendline_backward(trendline.backward_period);
        }

        if let Some(intercept) = trendline.intercept {
            // Write the c:intercept element.
            self.write_trendline_intercept(intercept);
        }

        if trendline.display_r_squared {
            // Write the c:dispRSqr element.
            self.write_disp_rsqr();
        }

        if trendline.display_equation {
            // Write the c:dispEq element.
            self.write_trendline_display_equation(trendline);
        }

        self.writer.xml_end_tag("c:trendline");
    }

    // Write the <c:name> element.
    fn write_trendline_name(&mut self, name: &str) {
        self.writer.xml_data_element_only("c:name", name);
    }

    // Write the <c:trendlineType> element.
    fn write_trendline_type(&mut self, trendline: &ChartTrendline) {
        let attributes = [("val", trendline.trend_type.to_string())];

        self.writer.xml_empty_tag("c:trendlineType", &attributes);
    }

    // Write the <c:forward> element.
    fn write_trendline_forward(&mut self, value: f64) {
        let attributes = [("val", value.to_string())];

        self.writer.xml_empty_tag("c:forward", &attributes);
    }

    // Write the <c:backward> element.
    fn write_trendline_backward(&mut self, value: f64) {
        let attributes = [("val", value.to_string())];

        self.writer.xml_empty_tag("c:backward", &attributes);
    }

    // Write the <c:dispRSqr> element.
    fn write_disp_rsqr(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:dispRSqr", &attributes);
    }

    // Write the <c:dispEq> element.
    fn write_trendline_display_equation(&mut self, trendline: &ChartTrendline) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:dispEq", &attributes);

        // Write the c:trendlineLbl element.
        self.write_trendline_label(trendline);
    }

    // Write the <c:trendlineLbl> element.
    fn write_trendline_label(&mut self, trendline: &ChartTrendline) {
        self.writer.xml_start_tag_only("c:trendlineLbl");

        // Write the c:layout element.
        self.write_layout();

        self.write_number_format("General", false);

        // Write the c:spPr formatting element.
        self.write_sp_pr(&trendline.label_format);

        // Write the trendline label font elements.
        if let Some(font) = &trendline.label_font {
            self.write_axis_font(font);
        }

        self.writer.xml_end_tag("c:trendlineLbl");
    }

    // Write the <c:period> element.
    fn write_trendline_period(&mut self, value: u8) {
        let attributes = [("val", value.to_string())];

        self.writer.xml_empty_tag("c:period", &attributes);
    }

    // Write the <c:intercept> element.
    fn write_trendline_intercept(&mut self, value: f64) {
        let attributes = [("val", value.to_string())];

        self.writer.xml_empty_tag("c:intercept", &attributes);
    }

    // Write the <c:errBars> element.
    fn write_error_bar(&mut self, axis: &str, error_bars: &ChartErrorBars) {
        self.writer.xml_start_tag_only("c:errBars");

        // Write the c:errDir element.
        self.write_error_bar_direction(axis);

        // Write the c:errBarType element.
        self.write_error_bar_type(error_bars.direction);

        // Write the c:errValType element.
        self.write_err_direction_type(&error_bars.error_type);

        if !error_bars.has_end_cap {
            // Write the c:noEndCap element.
            self.write_error_bar_no_end_cap();
        }

        match &error_bars.error_type {
            ChartErrorBarsType::FixedValue(value)
            | ChartErrorBarsType::Percentage(value)
            | ChartErrorBarsType::StandardDeviation(value) => {
                // Write the c:val element.
                self.write_error_value(*value);
            }
            ChartErrorBarsType::Custom(_, _) => self.write_custom_error_bar_values(error_bars),
            ChartErrorBarsType::StandardError => {}
        }

        // Write the c:spPr formatting element.
        self.write_sp_pr(&error_bars.format);

        self.writer.xml_end_tag("c:errBars");
    }

    // Write the <c:errDir> element.
    fn write_error_bar_direction(&mut self, axis: &str) {
        if !axis.is_empty() {
            let attributes = vec![("val", axis.to_string())];
            self.writer.xml_empty_tag("c:errDir", &attributes);
        }
    }

    // Write the <c:errBarType> element.
    fn write_error_bar_type(&mut self, direction: ChartErrorBarsDirection) {
        let attributes = vec![("val", direction.to_string())];

        self.writer.xml_empty_tag("c:errBarType", &attributes);
    }

    // Write the <c:errValType> element.
    fn write_err_direction_type(&mut self, bar_type: &ChartErrorBarsType) {
        let attributes = vec![("val", bar_type.to_string())];

        self.writer.xml_empty_tag("c:errValType", &attributes);
    }

    // Write the <c:noEndCap> element.
    fn write_error_bar_no_end_cap(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:noEndCap", &attributes);
    }

    // Write the <c:val> element.
    fn write_error_value(&mut self, value: f64) {
        let attributes = [("val", value.to_string())];

        self.writer.xml_empty_tag("c:val", &attributes);
    }

    // Write the custom error sub-elements
    fn write_custom_error_bar_values(&mut self, error_bars: &ChartErrorBars) {
        self.writer.xml_start_tag_only("c:plus");
        self.write_cache_ref(&error_bars.plus_range, true);
        self.writer.xml_end_tag("c:plus");

        self.writer.xml_start_tag_only("c:minus");
        self.write_cache_ref(&error_bars.minus_range, true);
        self.writer.xml_end_tag("c:minus");
    }

    // Write the <c:upDownBars> element.
    fn write_up_down_bars(&mut self) {
        self.writer.xml_start_tag_only("c:upDownBars");

        // Write the c:gapWidth element.
        self.write_gap_width(150);

        // Write the c:upBars element.
        self.write_up_bars();

        // Write the c:downBars element.
        self.write_down_bars();

        self.writer.xml_end_tag("c:upDownBars");
    }

    // Write the <c:upBars> element.
    fn write_up_bars(&mut self) {
        if self.up_bar_format.has_formatting() {
            self.writer.xml_start_tag_only("c:upBars");

            // Write the c:spPr element.
            self.write_sp_pr(&self.up_bar_format.clone());

            self.writer.xml_end_tag("c:upBars");
        } else {
            self.writer.xml_empty_tag_only("c:upBars");
        }
    }

    // Write the <c:downBars> element.
    fn write_down_bars(&mut self) {
        if self.down_bar_format.has_formatting() {
            self.writer.xml_start_tag_only("c:downBars");

            // Write the c:spPr element.
            self.write_sp_pr(&self.down_bar_format.clone());

            self.writer.xml_end_tag("c:downBars");
        } else {
            self.writer.xml_empty_tag_only("c:downBars");
        }
    }

    // Write the <c:hiLowLines> element.
    fn write_hi_low_lines(&mut self) {
        if self.high_low_lines_format.has_formatting() {
            self.writer.xml_start_tag_only("c:hiLowLines");

            // Write the c:spPr element.
            self.write_sp_pr(&self.high_low_lines_format.clone());

            self.writer.xml_end_tag("c:hiLowLines");
        } else {
            self.writer.xml_empty_tag_only("c:hiLowLines");
        }
    }

    // Write the <c:dropLines> element.
    fn write_drop_lines(&mut self) {
        if self.drop_lines_format.has_formatting() {
            self.writer.xml_start_tag_only("c:dropLines");

            // Write the c:spPr element.
            self.write_sp_pr(&self.drop_lines_format.clone());

            self.writer.xml_end_tag("c:dropLines");
        } else {
            self.writer.xml_empty_tag_only("c:dropLines");
        }
    }

    // Write the <c:dTable> element.
    fn write_data_table(&mut self, table: &ChartDataTable) {
        self.writer.xml_start_tag_only("c:dTable");

        // Write the c:showHorzBorder element.
        if table.show_horizontal_borders {
            self.write_show_horz_border();
        }

        // Write the c:showVertBorder element.
        if table.show_vertical_borders {
            self.write_show_vert_border();
        }

        // Write the c:showOutline element.
        if table.show_outline_borders {
            self.write_show_outline();
        }

        // Write the c:showKeys element.
        if table.show_legend_keys {
            self.write_show_keys();
        }

        // Write the c:spPr element.
        self.write_sp_pr(&table.format);

        // Write the trendline label font elements.
        if let Some(font) = &table.font {
            self.write_axis_font(font);
        }

        self.writer.xml_end_tag("c:dTable");
    }

    // Write the <c:showKeys> element.
    fn write_show_keys(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:showKeys", &attributes);
    }

    // Write the <c:showHorzBorder> element.
    fn write_show_horz_border(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:showHorzBorder", &attributes);
    }

    // Write the <c:showVertBorder> element.
    fn write_show_vert_border(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:showVertBorder", &attributes);
    }

    // Write the <c:showOutline> element.
    fn write_show_outline(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:showOutline", &attributes);
    }

    // Write the <c:showVal> element.
    fn write_show_val(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:showVal", &attributes);
    }

    // Write the <c:showCatName> element.
    fn write_show_category_name(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:showCatName", &attributes);
    }

    // Write the <c:showSerName> element.
    fn write_show_series_name(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:showSerName", &attributes);
    }

    // Write the <c:showPercent> element.
    fn write_show_percent(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:showPercent", &attributes);
    }

    // Write the <c:separator> element.
    fn write_separator(&mut self, separator: char) {
        self.writer
            .xml_data_element_only("c:separator", &format!("{separator} "));
    }

    // Write the <c:showLeaderLines> element for Excel 2007 (mainly only for Pie
    // and Doughnut).
    fn write_show_leader_lines_2007(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:showLeaderLines", &attributes);
    }

    // Write the <c:showLeaderLines> element for Excel 2015+ (mainly for charts
    // that aren't Pie or Doughnut).
    fn write_show_leader_lines_2015(&mut self) {
        let attributes = [
            ("uri", "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}"),
            (
                "xmlns:c15",
                "http://schemas.microsoft.com/office/drawing/2012/chart",
            ),
        ];

        self.writer.xml_start_tag_only("c:extLst");
        self.writer.xml_start_tag("c:ext", &attributes);

        self.writer
            .xml_empty_tag("c15:showLeaderLines", &[("val", "1")]);
        self.writer.xml_end_tag("c:ext");
        self.writer.xml_end_tag("c:extLst");
    }

    // Write the <c:showLegendKey> element.
    fn write_show_legend_key(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:showLegendKey", &attributes);
    }

    // Write the <c:dLblPos> element.
    fn write_d_lbl_pos(&mut self, position: ChartDataLabelPosition) {
        let attributes = [("val", position.to_string())];

        self.writer.xml_empty_tag("c:dLblPos", &attributes);
    }

    // Write the <c:delete> element.
    fn write_delete(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:delete", &attributes);
    }

    // Write the <c:symbol> element.
    fn write_symbol(&mut self, marker: &ChartMarker) {
        let mut attributes = vec![];

        if let Some(marker_type) = marker.marker_type {
            attributes.push(("val", marker_type.to_string()));
        } else if marker.none {
            attributes.push(("val", "none".to_string()));
        }

        self.writer.xml_empty_tag("c:symbol", &attributes);
    }

    // Write the <c:size> element.
    fn write_size(&mut self, size: u8) {
        let attributes = [("val", size.to_string())];

        self.writer.xml_empty_tag("c:size", &attributes);
    }

    // Write the <c:varyColors> element.
    fn write_vary_colors(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:varyColors", &attributes);
    }

    // Write the <c:firstSliceAng> element.
    fn write_first_slice_ang(&mut self) {
        let attributes = [("val", self.rotation.to_string())];

        self.writer.xml_empty_tag("c:firstSliceAng", &attributes);
    }

    // Write the <c:holeSize> element.
    fn write_hole_size(&mut self) {
        let attributes = [("val", self.hole_size.to_string())];

        self.writer.xml_empty_tag("c:holeSize", &attributes);
    }

    // Write the <c:txPr> element.
    fn write_axis_font(&mut self, font: &ChartFont) {
        self.writer.xml_start_tag_only("c:txPr");

        // Write the a:bodyPr element.
        self.write_a_body_pr(font, false);

        // Write the a:lstStyle element.
        self.write_a_lst_style();

        self.writer.xml_start_tag_only("a:p");

        // Write the a:pPr element.
        self.write_a_p_pr_rich(font);

        // Write the a:endParaRPr element.
        self.write_a_end_para_rpr();

        self.writer.xml_end_tag("a:p");

        self.writer.xml_end_tag("c:txPr");
    }

    // Write the <c:txPr> element.
    fn write_tx_pr(&mut self, font: &ChartFont, is_horizontal: bool) {
        self.writer.xml_start_tag_only("c:txPr");

        // Write the a:bodyPr element.
        self.write_a_body_pr(font, is_horizontal);

        // Write the a:lstStyle element.
        self.write_a_lst_style();

        // Write the a:p element.
        self.write_a_p_formula(font);

        self.writer.xml_end_tag("c:txPr");
    }

    // Write the <a:p> element.
    fn write_a_p_formula(&mut self, font: &ChartFont) {
        self.writer.xml_start_tag_only("a:p");

        // Write the a:pPr element.
        self.write_a_p_pr(font);

        // Write the a:endParaRPr element.
        self.write_a_end_para_rpr();

        self.writer.xml_end_tag("a:p");
    }

    // Write the <a:pPr> element.
    fn write_a_p_pr(&mut self, font: &ChartFont) {
        let mut attributes = vec![];

        if let Some(right_to_left) = font.right_to_left {
            attributes.push(("rtl", right_to_left.to_xml_bool()));
        }

        self.writer.xml_start_tag("a:pPr", &attributes);

        // Write the a:defRPr element.
        self.write_a_def_rpr(font);

        self.writer.xml_end_tag("a:pPr");
    }

    // Write the <a:bodyPr> element.
    fn write_a_body_pr(&mut self, font: &ChartFont, is_horizontal: bool) {
        let mut attributes = vec![];

        let rotation = match font.rotation {
            Some(rotation) => rotation,
            None => {
                // Handle defaults for vertical and horizontal rotations.
                if is_horizontal {
                    -90
                } else {
                    360 // To distinguish from user defined 0.
                }
            }
        };

        match rotation {
            360 => {}
            270 => {
                // Stacked text.
                attributes.push(("rot", "0".to_string()));
                attributes.push(("vert", "wordArtVert".to_string()));
            }
            271 => {
                // East Asian vertical.
                attributes.push(("rot", "0".to_string()));
                attributes.push(("vert", "eaVert".to_string()));
            }
            _ => {
                // Convert the rotation angle to Excel's units.
                let rotation = i32::from(rotation) * 60_000;
                attributes.push(("rot", rotation.to_string()));
                attributes.push(("vert", "horz".to_string()));
            }
        }

        self.writer.xml_empty_tag("a:bodyPr", &attributes);
    }

    // Write the <a:lstStyle> element.
    fn write_a_lst_style(&mut self) {
        self.writer.xml_empty_tag_only("a:lstStyle");
    }

    // Write the <a:defRPr> element.
    fn write_a_def_rpr(&mut self, font: &ChartFont) {
        self.write_font_elements("a:defRPr", font);
    }

    // Write the <a:rPr> element.
    fn write_a_r_pr(&mut self, font: &ChartFont) {
        self.write_font_elements("a:rPr", font);
    }

    // Write font sub-elements shared between <a:defRPr> and <a:rPr> elements.
    fn write_font_elements(&mut self, tag: &str, font: &ChartFont) {
        let mut attributes = vec![];

        if tag == "a:rPr" {
            attributes.push(("lang", "en-US".to_string()));
        }

        if font.size > 0.0 {
            attributes.push(("sz", font.size.to_string()));
        }

        if let Some(boolean) = font.bold {
            attributes.push(("b", boolean.to_xml_bool()));
        }

        if font.italic || (font.bold.is_some() && !font.has_default_bold) {
            attributes.push(("i", font.italic.to_xml_bool()));
        }

        if font.underline {
            attributes.push(("u", "sng".to_string()));
        }

        if font.has_baseline {
            attributes.push(("baseline", "0".to_string()));
        }

        if font.is_latin() || !font.color.is_auto_or_default() {
            self.writer.xml_start_tag(tag, &attributes);

            if !font.color.is_auto_or_default() {
                self.write_a_solid_fill(font.color, 0);
            }

            if font.is_latin() {
                // Write the a:latin element.
                self.write_a_latin(font);
            }

            self.writer.xml_end_tag(tag);
        } else {
            self.writer.xml_empty_tag(tag, &attributes);
        }
    }

    // Write the <a:latin> element.
    fn write_a_latin(&mut self, font: &ChartFont) {
        let mut attributes = vec![];

        if !font.name.is_empty() {
            attributes.push(("typeface", font.name.to_string()));
        }

        if font.pitch_family > 0 {
            attributes.push(("pitchFamily", font.pitch_family.to_string()));
        }

        if font.character_set > 0 || font.pitch_family > 0 {
            attributes.push(("charset", font.character_set.to_string()));
        }

        self.writer.xml_empty_tag("a:latin", &attributes);
    }

    // Write the <a:t> element.
    fn write_a_t(&mut self, name: &str) {
        self.writer.xml_data_element_only("a:t", name);
    }

    // Write the <a:endParaRPr> element.
    fn write_a_end_para_rpr(&mut self) {
        let attributes = [("lang", "en-US")];

        self.writer.xml_empty_tag("a:endParaRPr", &attributes);
    }

    // Write the <c:spPr> element.
    fn write_sp_pr(&mut self, format: &ChartFormat) {
        if !format.has_formatting() {
            return;
        }

        self.writer.xml_start_tag_only("c:spPr");

        if format.no_fill {
            self.writer.xml_empty_tag_only("a:noFill");
        } else if let Some(solid_fill) = &format.solid_fill {
            // Write the a:solidFill element.
            self.write_a_solid_fill(solid_fill.color, solid_fill.transparency);
        } else if let Some(pattern_fill) = &format.pattern_fill {
            // Write the a:pattFill element.
            self.write_a_patt_fill(pattern_fill);
        } else if let Some(gradient_fill) = &format.gradient_fill {
            // Write the a:gradFill element.
            self.write_gradient_fill(gradient_fill);
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
            // Round width to nearest 0.25, like Excel.
            let width = ((*width + 0.125) * 4.0).floor() / 4.0;

            // Convert to Excel internal units.
            let width = (12700.0 * width).ceil() as u32;

            attributes.push(("w", width.to_string()));
        }

        if line.color != Color::Default || line.dash_type != ChartLineDashType::Solid || line.hidden
        {
            self.writer.xml_start_tag("a:ln", &attributes);

            if line.hidden {
                // Write the a:noFill element.
                self.write_a_no_fill();
            } else {
                if line.color != Color::Default {
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
            self.writer.xml_empty_tag("a:ln", &attributes);
        }
    }

    // Write the <a:ln> element.
    fn write_a_ln_none(&mut self) {
        self.writer.xml_start_tag_only("a:ln");

        // Write the a:noFill element.
        self.write_a_no_fill();

        self.writer.xml_end_tag("a:ln");
    }

    // Write the <a:solidFill> element.
    fn write_a_solid_fill(&mut self, color: Color, transparency: u8) {
        self.writer.xml_start_tag_only("a:solidFill");

        // Write the color element.
        self.write_color(color, transparency);

        self.writer.xml_end_tag("a:solidFill");
    }

    // Write the <a:pattFill> element.
    fn write_a_patt_fill(&mut self, fill: &ChartPatternFill) {
        let attributes = [("prst", fill.pattern.to_string())];

        self.writer.xml_start_tag("a:pattFill", &attributes);

        if fill.foreground_color != Color::Default {
            // Write the <a:fgClr> element.
            self.writer.xml_start_tag_only("a:fgClr");
            self.write_color(fill.foreground_color, 0);
            self.writer.xml_end_tag("a:fgClr");
        }

        if fill.background_color != Color::Default {
            // Write the <a:bgClr> element.
            self.writer.xml_start_tag_only("a:bgClr");
            self.write_color(fill.background_color, 0);
            self.writer.xml_end_tag("a:bgClr");
        } else if fill.background_color == Color::Default && fill.foreground_color != Color::Default
        {
            // If there is a foreground color but no background color then we
            // need to write a default background color.
            self.writer.xml_start_tag_only("a:bgClr");
            self.write_color(Color::White, 0);
            self.writer.xml_end_tag("a:bgClr");
        }

        self.writer.xml_end_tag("a:pattFill");
    }

    // Write the <a:gradFill> element.
    fn write_gradient_fill(&mut self, fill: &ChartGradientFill) {
        let mut attributes = vec![];

        if fill.gradient_type != ChartGradientFillType::Linear {
            attributes.push(("flip", "none"));
            attributes.push(("rotWithShape", "1"));
        }

        self.writer.xml_start_tag("a:gradFill", &attributes);
        self.writer.xml_start_tag_only("a:gsLst");

        for gradient_stop in &fill.gradient_stops {
            // Write the a:gs element.
            self.write_gradient_stop(gradient_stop);
        }

        self.writer.xml_end_tag("a:gsLst");

        if fill.gradient_type == ChartGradientFillType::Linear {
            // Write the a:lin element.
            self.write_gradient_fill_angle(fill.angle);
        } else {
            // Write the a:path element.
            self.write_gradient_path(fill.gradient_type);
        }

        self.writer.xml_end_tag("a:gradFill");
    }

    // Write the <a:gs> element.
    fn write_gradient_stop(&mut self, gradient_stop: &ChartGradientStop) {
        let position = 1000 * u32::from(gradient_stop.position);
        let attributes = [("pos", position.to_string())];

        self.writer.xml_start_tag("a:gs", &attributes);
        self.write_color(gradient_stop.color, 0);

        self.writer.xml_end_tag("a:gs");
    }

    // Write the <a:lin> element.
    fn write_gradient_fill_angle(&mut self, angle: u16) {
        let angle = 60_000 * u32::from(angle);
        let attributes = [("ang", angle.to_string()), ("scaled", "0".to_string())];

        self.writer.xml_empty_tag("a:lin", &attributes);
    }

    // Write the <a:path> element.
    fn write_gradient_path(&mut self, gradient_type: ChartGradientFillType) {
        let mut attributes = vec![];

        match gradient_type {
            ChartGradientFillType::Radial => attributes.push(("path", "circle")),
            ChartGradientFillType::Rectangular => attributes.push(("path", "rect")),
            ChartGradientFillType::Path => attributes.push(("path", "shape")),
            ChartGradientFillType::Linear => {}
        }

        self.writer.xml_start_tag("a:path", &attributes);

        // Write the a:fillToRect element.
        self.write_a_fill_to_rect(gradient_type);

        self.writer.xml_end_tag("a:path");

        // Write the a:tileRect element.
        self.write_a_tile_rect(gradient_type);
    }

    // Write the <a:fillToRect> element.
    fn write_a_fill_to_rect(&mut self, gradient_type: ChartGradientFillType) {
        let mut attributes = vec![];

        match gradient_type {
            ChartGradientFillType::Path => {
                attributes.push(("l", "50000"));
                attributes.push(("t", "50000"));
                attributes.push(("r", "50000"));
                attributes.push(("b", "50000"));
            }
            _ => {
                attributes.push(("l", "100000"));
                attributes.push(("t", "100000"));
            }
        }

        self.writer.xml_empty_tag("a:fillToRect", &attributes);
    }

    // Write the <a:tileRect> element.
    fn write_a_tile_rect(&mut self, gradient_type: ChartGradientFillType) {
        let mut attributes = vec![];

        match gradient_type {
            ChartGradientFillType::Rectangular | ChartGradientFillType::Radial => {
                attributes.push(("r", "-100000"));
                attributes.push(("b", "-100000"));
            }
            _ => {}
        }

        self.writer.xml_empty_tag("a:tileRect", &attributes);
    }

    // Write the <a:srgbClr> element.
    fn write_color(&mut self, color: Color, transparency: u8) {
        match color {
            Color::Theme(_, _) => {
                let (scheme, lum_mod, lum_off) = color.chart_scheme();
                if !scheme.is_empty() {
                    // Write the a:schemeClr element.
                    self.write_a_scheme_clr(scheme, lum_mod, lum_off, transparency);
                }
            }
            Color::Automatic => {
                let attributes = [("val", "window"), ("lastClr", "FFFFFF")];

                self.writer.xml_empty_tag("a:sysClr", &attributes);
            }
            _ => {
                let attributes = [("val", color.rgb_hex_value())];

                if transparency > 0 {
                    self.writer.xml_start_tag("a:srgbClr", &attributes);

                    // Write the a:alpha element.
                    self.write_a_alpha(transparency);

                    self.writer.xml_end_tag("a:srgbClr");
                } else {
                    self.writer.xml_empty_tag("a:srgbClr", &attributes);
                }
            }
        }
    }

    // Write the <a:schemeClr> element.
    fn write_a_scheme_clr(&mut self, scheme: String, lum_mod: u32, lum_off: u32, transparency: u8) {
        let attributes = [("val", scheme)];

        if lum_mod > 0 || lum_off > 0 || transparency > 0 {
            self.writer.xml_start_tag("a:schemeClr", &attributes);

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
            self.writer.xml_empty_tag("a:schemeClr", &attributes);
        }
    }

    // Write the <a:lumMod> element.
    fn write_a_lum_mod(&mut self, lum_mod: u32) {
        let attributes = [("val", lum_mod.to_string())];

        self.writer.xml_empty_tag("a:lumMod", &attributes);
    }

    // Write the <a:lumOff> element.
    fn write_a_lum_off(&mut self, lum_off: u32) {
        let attributes = [("val", lum_off.to_string())];

        self.writer.xml_empty_tag("a:lumOff", &attributes);
    }

    // Write the <a:alpha> element.
    fn write_a_alpha(&mut self, transparency: u8) {
        let transparency = u32::from(100 - transparency) * 1000;

        let attributes = [("val", transparency.to_string())];

        self.writer.xml_empty_tag("a:alpha", &attributes);
    }

    // Write the <a:noFill> element.
    fn write_a_no_fill(&mut self) {
        self.writer.xml_empty_tag_only("a:noFill");
    }

    // Write the <a:prstDash> element.
    fn write_a_prst_dash(&mut self, line: &ChartLine) {
        let attributes = [("val", line.dash_type.to_string())];

        self.writer.xml_empty_tag("a:prstDash", &attributes);
    }

    // Write the <c:radarStyle> element.
    fn write_radar_style(&mut self) {
        let mut attributes = vec![];

        if self.chart_type == ChartType::RadarFilled {
            attributes.push(("val", "filled".to_string()));
        } else {
            attributes.push(("val", "marker".to_string()));
        }

        self.writer.xml_empty_tag("c:radarStyle", &attributes);
    }

    // Write the <c:majorTickMark> element.
    fn write_major_tick_mark(&mut self, position: ChartAxisTickType) {
        let attributes = [("val", position.to_string())];

        self.writer.xml_empty_tag("c:majorTickMark", &attributes);
    }

    // Write the <c:minorTickMark> element.
    fn write_minor_tick_mark(&mut self, tick_type: ChartAxisTickType) {
        let attributes = [("val", tick_type.to_string())];

        self.writer.xml_empty_tag("c:minorTickMark", &attributes);
    }

    // Write the <c:gapWidth> element.
    fn write_gap_width(&mut self, gap: u16) {
        let attributes = [("val", gap.to_string())];

        self.writer.xml_empty_tag("c:gapWidth", &attributes);
    }

    // Write the <c:overlap> element.
    fn write_overlap(&mut self) {
        let attributes = [("val", self.overlap.to_string())];

        self.writer.xml_empty_tag("c:overlap", &attributes);
    }

    // Write the <c:smooth> element.
    fn write_smooth(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:smooth", &attributes);
    }

    // Write the <c:style> element.
    fn write_style(&mut self) {
        let attributes = [("val", self.style.to_string())];

        self.writer.xml_empty_tag("c:style", &attributes);
    }

    // Write the <c:autoTitleDeleted> element.
    fn write_auto_title_deleted(&mut self) {
        let attributes = [("val", "1")];

        self.writer.xml_empty_tag("c:autoTitleDeleted", &attributes);
    }

    // Write the <c:title> element.
    fn write_title_formula(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag_only("c:title");

        // Write the c:tx element.
        self.write_tx_formula(title);

        // Write the c:layout element.
        self.write_layout();

        if title.format.has_formatting() {
            // Write the c:spPr formatting element.
            self.write_sp_pr(&title.format.clone());
        } else {
            // Write the c:txPr element.
            self.write_tx_pr(&title.font, title.is_horizontal);
        }

        self.writer.xml_end_tag("c:title");
    }

    // Write the <c:tx> element.
    fn write_tx_formula(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag_only("c:tx");

        // Title is always a string type.
        self.write_str_ref(&title.range);

        self.writer.xml_end_tag("c:tx");
    }

    // Write the <c:title> element.
    fn write_title_rich(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag_only("c:title");

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
        self.writer.xml_start_tag_only("c:title");

        // Write the c:layout element.
        self.write_layout();

        // Write the c:spPr element.
        self.write_sp_pr(&title.format.clone());

        self.writer.xml_end_tag("c:title");
    }

    // Write the <c:tx> element.
    fn write_tx_rich(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag_only("c:tx");

        // Write the c:rich element.
        self.write_rich(title);

        self.writer.xml_end_tag("c:tx");
    }

    // Write the <c:tx> element.
    fn write_tx_value(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag_only("c:tx");

        self.writer.xml_data_element_only("c:v", &title.name);

        self.writer.xml_end_tag("c:tx");
    }

    // Write the <c:rich> element.
    fn write_rich(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag_only("c:rich");

        // Write the a:bodyPr element.
        self.write_a_body_pr(&title.font, title.is_horizontal);

        // Write the a:lstStyle element.
        self.write_a_lst_style();

        // Write the a:p element.
        self.write_a_p_rich(title);

        self.writer.xml_end_tag("c:rich");
    }

    // Write the <a:p> element.
    fn write_a_p_rich(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag_only("a:p");

        if !title.ignore_rich_para {
            // Write the a:pPr element.
            self.write_a_p_pr_rich(&title.font);
        }

        // Write the a:r element.
        self.write_a_r(title);

        self.writer.xml_end_tag("a:p");
    }

    // Write the <a:pPr> element.
    fn write_a_p_pr_rich(&mut self, font: &ChartFont) {
        let mut attributes = vec![];

        if let Some(right_to_left) = font.right_to_left {
            attributes.push(("rtl", right_to_left.to_xml_bool()));
        }

        self.writer.xml_start_tag("a:pPr", &attributes);

        // Write the a:defRPr element.
        self.write_a_def_rpr(font);

        self.writer.xml_end_tag("a:pPr");
    }

    // Write the <a:r> element.
    fn write_a_r(&mut self, title: &ChartTitle) {
        self.writer.xml_start_tag_only("a:r");

        // Write the a:rPr element.
        self.write_a_r_pr(&title.font);

        // Write the a:t element.
        self.write_a_t(&title.name);

        self.writer.xml_end_tag("a:r");
    }

    // Write the <c:dispBlanksAs> element.
    fn write_disp_blanks_as(&mut self) {
        if let Some(show_empty_cells) = self.show_empty_cells_as {
            let attributes = [("val", show_empty_cells.to_string())];

            self.writer.xml_empty_tag("c:dispBlanksAs", &attributes);
        }
    }

    // Write the <dispNaAsBlank> element. This is an Excel 16 extension.
    fn write_disp_na_as_blank(&mut self) {
        let attributes = [
            ("uri", "{56B9EC1D-385E-4148-901F-78D8002777C0}"),
            (
                "xmlns:c16r3",
                "http://schemas.microsoft.com/office/drawing/2017/03/chart",
            ),
        ];

        self.writer.xml_start_tag_only("c:extLst");
        self.writer.xml_start_tag("c:ext", &attributes);
        self.writer.xml_start_tag_only("c16r3:dataDisplayOptions16");

        self.writer
            .xml_empty_tag("c16r3:dispNaAsBlank", &[("val", "1")]);

        self.writer.xml_end_tag("c16r3:dataDisplayOptions16");
        self.writer.xml_end_tag("c:ext");
        self.writer.xml_end_tag("c:extLst");
    }
}

// -----------------------------------------------------------------------
// Traits.
// -----------------------------------------------------------------------

/// Trait to map types into an `ChartRange`.
///
/// The 2 most common types of range used in `rust_xlsxwriter` charts are:
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

impl IntoChartRange for &ChartRange {
    fn new_chart_range(&self) -> ChartRange {
        (*self).clone()
    }
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

impl IntoChartFormat for &mut ChartGradientFill {
    fn new_chart_format(&self) -> ChartFormat {
        ChartFormat::new().set_gradient_fill(self).clone()
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

    fn name(&self) -> String {
        self.name.clone()
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
// Secondary structs and enums
// -----------------------------------------------------------------------

// -----------------------------------------------------------------------
// ChartSeries
// -----------------------------------------------------------------------

/// The `ChartSeries` struct represents a chart series.
///
/// A chart in Excel can contain one of more data series. The `ChartSeries`
/// struct represents the Category and Value ranges, and the formatting and
/// options for the chart series.
///
/// It is used in conjunction with the [`Chart`] struct.
///
/// # Examples
///
/// A simple chart example using the `rust_xlsxwriter` library.
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
    pub(crate) title: ChartTitle,
    pub(crate) format: ChartFormat,
    pub(crate) marker: Option<ChartMarker>,
    pub(crate) data_label: Option<ChartDataLabel>,
    pub(crate) custom_data_labels: Vec<ChartDataLabel>,
    pub(crate) points: Vec<ChartPoint>,
    pub(crate) gap: u16,
    pub(crate) overlap: i8,
    pub(crate) invert_if_negative: bool,
    pub(crate) inverted_color: Color,
    pub(crate) trendline: ChartTrendline,
    pub(crate) x_error_bars: Option<ChartErrorBars>,
    pub(crate) y_error_bars: Option<ChartErrorBars>,
    pub(crate) delete_from_legend: bool,
    pub(crate) smooth: Option<bool>,
}

#[allow(clippy::new_without_default)]
impl ChartSeries {
    /// Create a new chart series object.
    ///
    /// Create a new chart series object. A chart in Excel must contain at least
    /// one data series. The `ChartSeries` struct represents the category and
    /// value ranges, and the formatting and options for the chart series.
    ///
    /// It is used in conjunction with the [`Chart`] struct.
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
    /// #
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
    /// #
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
            value_range: ChartRange::default(),
            category_range: ChartRange::default(),
            title: ChartTitle::new(),
            format: ChartFormat::default(),
            marker: None,
            data_label: None,
            points: vec![],
            custom_data_labels: vec![],
            gap: 150,
            overlap: 0,
            invert_if_negative: false,
            inverted_color: Color::Default,
            trendline: ChartTrendline::new(),
            x_error_bars: None,
            y_error_bars: None,
            delete_from_legend: false,
            smooth: None,
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
    /// # Parameters
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
    /// # Parameters
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
    /// # Parameters
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
    /// - [`ChartFormat::set_solid_fill()`]: Set the [`ChartSolidFill`] properties.
    /// - [`ChartFormat::set_pattern_fill()`]: Set the [`ChartPatternFill`] properties.
    /// - [`ChartFormat::set_gradient_fill()`]: Set the [`ChartGradientFill`] properties.
    /// - [`ChartFormat::set_no_fill()`]: Turn off the fill for the chart object.
    /// - [`ChartFormat::set_line()`]: Set the [`ChartLine`] properties.
    /// - [`ChartFormat::set_border()`]: Set the [`ChartBorder`] properties.
    ///   A synonym for [`ChartLine`] depending on context.
    /// - [`ChartFormat::set_no_line()`]: Turn off the line for the chart object.
    /// - [`ChartFormat::set_no_border()`]: Turn off the border for the chart object.
    ///
    /// # Parameters
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
    /// # Parameters
    ///
    /// `marker`: A [`ChartMarker`] instance.
    ///
    /// # Examples
    ///
    /// An example of adding markers to a Line chart.
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker.png">
    ///
    pub fn set_marker(&mut self, marker: &ChartMarker) -> &mut ChartSeries {
        self.marker = Some(marker.clone());
        self
    }

    /// Set the data labels for a chart series.
    ///
    /// Set the data labels and marker properties for a data series using a
    /// [`ChartDataLabel`] instance.
    ///
    /// # Parameters
    ///
    /// `data_label`: A [`ChartDataLabel`] instance.
    ///
    /// # Examples
    ///
    /// An example of adding data labels to a chart series.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_data_labels.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError, ChartDataLabel};
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
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_data_label(ChartDataLabel::new().show_value());
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_data_labels.png">
    ///
    pub fn set_data_label(&mut self, data_label: &ChartDataLabel) -> &mut ChartSeries {
        self.data_label = Some(data_label.clone());
        self
    }

    /// Set custom data labels for a data series.
    ///
    /// The [`set_data_label()`](ChartSeries::set_data_label) sets the data
    /// label properties for every label in a series but it is occasionally
    /// required to set separate properties for individual data labels, or set a
    /// custom display value, or format or hide some of the labels. This can be
    /// achieved with the `set_custom_data_labels()` method, see the examples
    /// below.
    ///
    /// # Parameters
    ///
    /// `data_labels`: A slice of [`ChartDataLabel`] objects.
    ///
    /// # Examples
    ///
    /// An example of adding custom data labels to a chart series. This is
    /// useful when you want to label the points of a data series with
    /// information that isn't contained in the value or category names.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_custom_data_labels1.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartDataLabel, ChartType, Workbook, XlsxError};
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
    ///     // Create some custom data labels.
    ///     let data_labels = [
    ///         ChartDataLabel::new().set_value("Alice").to_custom(),
    ///         ChartDataLabel::new().set_value("Bob").to_custom(),
    ///         ChartDataLabel::new().set_value("Carol").to_custom(),
    ///         ChartDataLabel::new().set_value("Dave").to_custom(),
    ///         ChartDataLabel::new().set_value("Eve").to_custom(),
    ///         ChartDataLabel::new().set_value("Frank").to_custom(),
    ///     ];
    ///
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_custom_data_labels(&data_labels);
    ///
    ///     // Turn legend off for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_set_custom_data_labels1.png">
    ///
    /// This example shows how to get the data from cells. In Excel this is a
    /// single command called "Value from Cells" but in `rust_xlsxwriter` it
    /// needs to be broken down into a cell reference for each data label.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_custom_data_labels2.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartDataLabel, ChartType, Workbook, XlsxError};
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
    /// #     worksheet.write(0, 1, "Asia")?;
    /// #     worksheet.write(1, 1, "Africa")?;
    /// #     worksheet.write(2, 1, "Europe")?;
    /// #     worksheet.write(3, 1, "Americas")?;
    /// #     worksheet.write(4, 1, "Oceania")?;
    /// #     worksheet.write(5, 1, "Antarctic")?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Create some custom data labels.
    ///     let data_labels = [
    ///         ChartDataLabel::new().set_value("=Sheet1!$B$1").to_custom(),
    ///         ChartDataLabel::new().set_value("=Sheet1!$B$2").to_custom(),
    ///         ChartDataLabel::new().set_value("=Sheet1!$B$3").to_custom(),
    ///         ChartDataLabel::new().set_value("=Sheet1!$B$4").to_custom(),
    ///         ChartDataLabel::new().set_value("=Sheet1!$B$5").to_custom(),
    ///         ChartDataLabel::new().set_value("=Sheet1!$B$6").to_custom(),
    ///     ];
    ///
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_custom_data_labels(&data_labels);
    ///
    ///     // Turn legend off for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_set_custom_data_labels2.png">
    ///
    /// This example shows how to add default/non-custom data labels along with
    /// custom data labels. This is done in two ways: with an explicit
    /// `default()` data label and with an implicit default for points that
    /// aren't covered at the end of the list.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_custom_data_labels3.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartDataLabel, ChartType, Workbook, XlsxError};
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
    ///     // Create some custom data labels.
    ///     let data_labels = [
    ///         ChartDataLabel::default(),
    ///         ChartDataLabel::default(),
    ///         ChartDataLabel::new().set_value("Alice").to_custom(),
    ///         ChartDataLabel::new().set_value("Bob").to_custom(),
    ///         // All other points after this will get a default label.
    ///     ];
    ///
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_custom_data_labels(&data_labels);
    ///
    ///     // Turn legend off for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_set_custom_data_labels3.png">
    ///
    /// This example shows how to hide some of the data labels and keep others
    /// visible.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_custom_data_labels4.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartDataLabel, ChartType, Workbook, XlsxError};
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
    ///     // Create some custom data labels.
    ///     let data_labels = [
    ///         ChartDataLabel::default(),
    ///         ChartDataLabel::new().set_hidden().to_custom(),
    ///         ChartDataLabel::new().set_hidden().to_custom(),
    ///         ChartDataLabel::new().set_hidden().to_custom(),
    ///         ChartDataLabel::new().set_hidden().to_custom(),
    ///         ChartDataLabel::default(),
    ///     ];
    ///
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_custom_data_labels(&data_labels);
    ///
    ///     // Turn legend off for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_set_custom_data_labels4.png">
    ///
    /// This example shows how to format some of the data labels and leave the
    /// rest with the default formatting.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_custom_data_labels5.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartDataLabel, ChartFormat, ChartLine, ChartSolidFill, ChartType, Workbook, XlsxError,
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
    ///     // Create some custom data labels.
    ///     let data_labels = [
    ///         ChartDataLabel::new()
    ///             .set_value("Start")
    ///             .set_format(
    ///                 ChartFormat::new()
    ///                     .set_border(ChartLine::new().set_color("#FF0000"))
    ///                     .set_solid_fill(ChartSolidFill::new().set_color("#FFFF00")),
    ///             )
    ///             .to_custom(),
    ///         ChartDataLabel::new().set_hidden().to_custom(),
    ///         ChartDataLabel::new().set_hidden().to_custom(),
    ///         ChartDataLabel::new().set_hidden().to_custom(),
    ///         ChartDataLabel::new().set_hidden().to_custom(),
    ///         ChartDataLabel::new()
    ///             .set_value("End")
    ///             .set_format(
    ///                 ChartFormat::new()
    ///                     .set_border(ChartLine::new().set_color("#FF0000"))
    ///                     .set_solid_fill(ChartSolidFill::new().set_color("#FFFF00")),
    ///             )
    ///             .to_custom(),
    ///     ];
    ///
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_custom_data_labels(&data_labels);
    ///
    ///     // Turn legend off for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_set_custom_data_labels5.png">
    ///
    pub fn set_custom_data_labels(&mut self, data_labels: &[ChartDataLabel]) -> &mut ChartSeries {
        if self.data_label.is_none() {
            self.data_label = Some(ChartDataLabel::default());
        }

        self.custom_data_labels = data_labels.to_vec();

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
    /// # Parameters
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
    /// #     Chart, ChartFormat, ChartPoint, ChartSolidFill, Workbook, XlsxError,
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
    ///     let mut chart = Chart::new_pie();
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$3")
    ///         .set_points(&points);
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// As a syntactic shortcut the `set_point_colors()` method allows you to set
    /// the colors of chart points with a simpler interface.
    ///
    /// Compare the example below with the previous more general example which
    /// both produce the same result.
    ///
    /// # Parameters
    ///
    /// `colors`: a slice of [`Color`] enum values or types that will
    /// convert into [`Color`] via [`IntoColor`].
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
    /// # use rust_xlsxwriter::{Chart, Workbook, XlsxError};
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
    ///     let mut chart = Chart::new_pie();
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$3")
    ///         .set_point_colors(&["#FF000", "#FFC000", "#FFFF00"]);
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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

    /// Set the trendline for a chart series.
    ///
    /// Excel allows you to add a trendline to a data series that represents the
    /// trend or regression of the data using different types of fit. A
    /// [`ChartTrendline`] struct reference is used to represents the options of
    /// Excel trendlines and can be added to a series via the
    /// [`ChartSeries::set_trendline()`] method.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/trendline_options.png">
    ///
    /// # Parameters
    ///
    /// `trendline`: A [`ChartTrendline`] reference.
    ///
    /// # Examples
    ///
    /// An example of adding a trendline to a chart data series. The options
    /// used are shown in the image above.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_trendline_intro.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartTrendline, ChartTrendlineType, ChartType, Workbook, XlsxError};
    /// #
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add some data for the chart.
    ///     worksheet.write(0, 0, 11.1)?;
    ///     worksheet.write(1, 0, 18.8)?;
    ///     worksheet.write(2, 0, 33.2)?;
    ///     worksheet.write(3, 0, 37.5)?;
    ///     worksheet.write(4, 0, 52.1)?;
    ///     worksheet.write(5, 0, 58.9)?;
    ///
    ///     // Create a trendline.
    ///     let mut trendline = ChartTrendline::new();
    ///     trendline
    ///         .set_type(ChartTrendlineType::Linear)
    ///         .display_equation(true)
    ///         .display_r_squared(true);
    ///
    ///     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with a trendline.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_trendline(&trendline);
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_trendline_intro.png">
    ///
    pub fn set_trendline(&mut self, trendline: &ChartTrendline) -> &mut ChartSeries {
        self.trendline = trendline.clone();
        self
    }

    /// Set the vertical error bars for a chart series.
    ///
    /// Error bars on Excel charts allow you to show margins of error for a series
    /// based on measures such as Standard Deviation, Standard Error, Fixed values,
    /// Percentages or even custom defined ranges.
    ///
    /// The `ChartErrorBars` struct represents the error bars for a chart series.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_error_bars_options.png">
    ///
    /// # Parameters
    ///
    /// `error_bars`: A [`ChartErrorBars`] reference.
    ///
    /// # Examples
    ///
    /// An example of adding error bars to a chart data series.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_error_bars_intro.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartErrorBars, ChartErrorBarsType, ChartLine, ChartType, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 11.1)?;
    /// #     worksheet.write(1, 0, 18.8)?;
    /// #     worksheet.write(2, 0, 33.2)?;
    /// #     worksheet.write(3, 0, 37.5)?;
    /// #     worksheet.write(4, 0, 52.1)?;
    /// #     worksheet.write(5, 0, 58.9)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with error bars.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_y_error_bars(
    ///             ChartErrorBars::new()
    ///                 .set_type(ChartErrorBarsType::StandardError)
    ///                 .set_format(ChartLine::new().set_color("#FF0000")),
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
    /// src="https://rustxlsxwriter.github.io/images/chart_error_bars_intro.png">
    ///
    ///
    ///
    pub fn set_y_error_bars(&mut self, error_bars: &ChartErrorBars) -> &mut ChartSeries {
        self.y_error_bars = Some(error_bars.clone());
        self
    }

    /// Set the horizontal error bars for a chart series.
    ///
    /// See [`ChartSeries::set_y_error_bars()`] above for a description of error
    /// bars and their properties.
    ///
    /// Horizontal error bars can only be set in Excel for Scatter and Bar charts.
    ///
    /// # Parameters
    ///
    /// `error_bars`: A [`ChartErrorBars`] reference.
    ///
    pub fn set_x_error_bars(&mut self, error_bars: &ChartErrorBars) -> &mut ChartSeries {
        self.x_error_bars = Some(error_bars.clone());
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
    /// # Parameters
    ///
    /// * `overlap`: Overlap percentage of columns in Bar/Column charts. The
    /// range is -100 <= overlap <= 100 and the default is 0.
    ///
    /// # Examples
    ///
    /// an example of setting the chart series gap and overlap. Note that it only
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
    /// # Parameters
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

    /// Set line type charts to smooth for a series.
    ///
    /// Line and Scatter charts can have a linear or smoothed line connecting
    /// their data points. Some chart types such as [`ChartType::ScatterSmooth`] have
    /// smoothed series by default and some such as
    /// [`ChartType::ScatterStraight`] don't.
    ///
    /// The `ChartSeries::set_smooth()` method can be used to turn on/off the
    /// smooth property for a series.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. The default depends on the chart
    ///   type.
    ///
    pub fn set_smooth(&mut self, enable: bool) -> &mut ChartSeries {
        self.smooth = Some(enable);
        self
    }

    /// Invert the color for negative values in a column/bar chart series.
    ///
    /// Bar and Column charts in Excel offer a series property called "Invert if
    /// negative". This isn't available for other types of charts.
    ///
    /// The negative values are shown as a white solid fill with a black border.
    /// To set the color of the negative part of the bar/column see
    /// [`set_invert_if_negative_color()`](ChartSeries::set_invert_if_negative_color)
    /// below.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the "Invert if negative" property
    /// for a chart series.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_series_set_invert_if_negative.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, -5)?;
    /// #     worksheet.write(2, 0, 20)?;
    /// #     worksheet.write(3, 0, -15)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series and set the "Invert if negative" property.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$5")
    ///         .set_invert_if_negative();
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_series_set_invert_if_negative.png">
    ///
    pub fn set_invert_if_negative(&mut self) -> &mut ChartSeries {
        self.invert_if_negative = true;
        self
    }

    /// Set the inverted color for negative values in a column/bar chart series.
    ///
    /// Bar and Column charts in Excel offer a series property called "Invert if
    /// negative" (see
    /// [`set_invert_if_negative()`](ChartSeries::set_invert_if_negative)
    /// above).
    ///
    /// The negative values are usually shown as a white solid fill with a black
    /// border but the `set_invert_if_negative_color()` method can be use to set
    /// a user defined color. This also requires that you set a
    /// [`ChartSolidFill`] for the series.
    ///
    /// # Parameters
    ///
    /// * `color` - The inverse color property defined by a [`Color`] enum
    ///   value.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the "Invert if negative" property and
    /// associated color for a chart series. This also requires that you set a solid
    /// fill color for the series.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_series_set_invert_if_negative_color.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartSolidFill, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, -5)?;
    /// #     worksheet.write(2, 0, 20)?;
    /// #     worksheet.write(3, 0, -15)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series and set the "Invert if negative" property and color.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$5")
    ///         .set_format(ChartSolidFill::new().set_color("#4C9900"))
    ///         .set_invert_if_negative_color("#FF6666");
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_series_set_invert_if_negative_color.png">
    ///
    pub fn set_invert_if_negative_color<T>(&mut self, color: T) -> &mut ChartSeries
    where
        T: IntoColor,
    {
        let color = color.new_color();
        if color.is_valid() {
            self.invert_if_negative = true;
            self.inverted_color = color;
        }
        self
    }

    /// Delete/hide the series name from the chart legend.
    ///
    /// The `delete_from_legend()` method deletes/hides the series name from the
    /// chart legend. This is sometimes required if there are a lot of secondary
    /// series and their names are cluttering the chart legend.
    ///
    /// Note, to hide all the names in the chart legend you should use the
    /// [`ChartLegend::set_hidden()`] method instead.
    ///
    /// See also the [`ChartTrendline::delete_from_legend()`] and the
    /// [`ChartLegend::delete_entries()`] methods.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating deleting/hiding a series name from the
    /// chart legend.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_series_delete_from_legend.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 30)?;
    /// #     worksheet.write(1, 0, 20)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(0, 1, 10)?;
    /// #     worksheet.write(1, 1, 10)?;
    /// #     worksheet.write(2, 1, 10)?;
    /// #     worksheet.write(0, 2, 20)?;
    /// #     worksheet.write(1, 2, 15)?;
    /// #     worksheet.write(2, 2, 30)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a series whose name will appear in the legend.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Add a series but delete/hide its names from the legend.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$B$1:$B$3")
    ///         .delete_from_legend(true);
    ///
    ///     // Add a series whose name will appear in the legend.
    ///     chart.add_series().set_values("Sheet1!$C$1:$C$3");
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 3, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_series_delete_from_legend.png">
    ///
    ///
    /// The default display without deleting the names from the legend would
    /// look like this:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_series_delete_from_legend2.png">
    ///
    pub fn delete_from_legend(&mut self, enable: bool) -> &mut ChartSeries {
        self.delete_from_legend = enable;
        self
    }
}

// -----------------------------------------------------------------------
// ChartRange
// -----------------------------------------------------------------------

#[derive(Clone, PartialEq)]
/// The `ChartRange` struct represents a chart range.
///
/// A struct to represent a chart range like `"Sheet1!$A$1:$A$4"`. The struct is
/// public to allow for the [`IntoChartRange`] trait and for a limited set of
/// edge cases, but it isn't generally required to be manipulated by the end
/// user.
///
/// It is used in conjunction with the [`Chart`] struct.
///
pub struct ChartRange {
    sheet_name: String,
    first_row: RowNum,
    first_col: ColNum,
    last_row: RowNum,
    last_col: ColNum,
    range_string: String,
    pub(crate) cache: ChartRangeCacheData,
}

impl Default for ChartRange {
    fn default() -> Self {
        Self::new_from_range("", 0, 0, 0, 0)
    }
}

impl ChartRange {
    /// Create a new `ChartRange` from a worksheet 5 tuple.
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating a new chart range.
    ///
    /// ```
    /// # // This code is available in examples/doc_chartrange_new_from_range.rs
    /// #
    /// # use rust_xlsxwriter::ChartRange;
    /// #
    /// # #[allow(unused_variables)]
    /// # fn main() {
    ///     // Same as "Sheet1!$A$1:$A$5".
    ///     let range = ChartRange::new_from_range("Sheet1", 0, 0, 4, 0);
    /// # }
    /// ```
    ///
    pub fn new_from_range(
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
            range_string: String::new(),
            cache: ChartRangeCacheData::new(),
        }
    }

    /// Create a new `ChartRange` from an Excel range formula.
    ///
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating a new chart range.
    ///
    ///
    /// ```
    /// # // This code is available in examples/doc_chartrange_new_from_string.rs
    /// #
    /// # use rust_xlsxwriter::ChartRange;
    /// #
    /// # #[allow(unused_variables)]
    /// # fn main() {
    ///     let range = ChartRange::new_from_string("Sheet1!$A$1:$A$5");
    /// # }
    /// ```
    ///
    pub fn new_from_string(range_string: &str) -> ChartRange {
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
            first_col = utility::column_name_to_number(caps.get(2).unwrap().as_str());
            last_col = utility::column_name_to_number(caps.get(4).unwrap().as_str());
        } else if let Some(caps) = CHART_CELL.captures(range_string) {
            sheet_name = caps.get(1).unwrap().as_str();
            first_row = caps.get(3).unwrap().as_str().parse::<u32>().unwrap() - 1;
            first_col = utility::column_name_to_number(caps.get(2).unwrap().as_str());
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
            cache: ChartRangeCacheData::new(),
        }
    }

    // Convert the row/col range into a chart range string.
    pub(crate) fn formula(&self) -> String {
        utility::chart_range(
            &self.sheet_name,
            self.first_row,
            self.first_col,
            self.last_row,
            self.last_col,
        )
    }

    // Convert the row/col range into an absolute chart range string.
    pub(crate) fn formula_abs(&self) -> String {
        utility::chart_range_abs(
            &self.sheet_name,
            self.first_row,
            self.first_col,
            self.last_row,
            self.last_col,
        )
    }

    // Convert the row/col range into a range error string.
    pub(crate) fn error_range(&self) -> String {
        utility::chart_error_range(
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

    // Get the number of X or Y data points in the range.
    pub(crate) fn number_of_points(&self) -> usize {
        let row_range = (self.last_row - self.first_row + 1) as usize;
        let col_range = (self.last_col - self.first_col + 1) as usize;

        std::cmp::max(row_range, col_range)
    }

    // Get the number of X and Y data points in the range.
    pub(crate) fn number_of_range_points(&self) -> (usize, usize) {
        let row_range = (self.last_row - self.first_row + 1) as usize;
        let col_range = (self.last_col - self.first_col + 1) as usize;

        (row_range, col_range)
    }

    // Set the start point in a 2D range. This is used to start incremental
    // ranges, see below.
    pub(crate) fn set_baseline(&mut self, row_order: bool) {
        if row_order {
            self.last_row = self.first_row;
        } else {
            self.last_col = self.first_col;
        }
    }

    // Increment a 1D slice in a 2D range. Used to generate sequential cell
    // ranges.
    pub(crate) fn increment(&mut self, row_order: bool) {
        if row_order {
            self.first_row += 1;
            self.last_row = self.first_row;
        } else {
            self.first_col += 1;
            self.last_col = self.first_col;
        }
    }

    // Check that the row/column values in the range are valid.
    pub(crate) fn validate(&self) -> Result<(), XlsxError> {
        let range = self.error_range();

        let error_message = format!("Sheet name error for range: '{range}'");
        utility::validate_sheetname(&self.sheet_name, &error_message)?;

        if self.first_row > self.last_row {
            return Err(XlsxError::ChartError(format!(
                "Range '{range}' has a first row greater than the last row"
            )));
        }

        if self.first_col > self.last_col {
            return Err(XlsxError::ChartError(format!(
                "Range '{range}' has a first column greater than the last column"
            )));
        }

        if self.first_row >= ROW_MAX || self.last_row >= ROW_MAX {
            return Err(XlsxError::ChartError(format!(
                "Range '{range}' has a first row greater than Excel limit of 1048576"
            )));
        }

        if self.first_col >= COL_MAX || self.last_col >= COL_MAX {
            return Err(XlsxError::ChartError(format!(
                "Range '{range}' has a first column greater than Excel limit of XFD/16384"
            )));
        }

        Ok(())
    }

    // Check that the range is 1D.
    pub(crate) fn is_1d(&self) -> bool {
        self.last_row - self.first_row == 0 || self.last_col - self.first_col == 0
    }

    /// Add data to the `ChartRange` cache.
    ///
    /// This method is only used to populate the chart data caches in test code.
    /// Outside of tests the library reads and populates the cache automatically.
    ///
    /// # Parameters
    ///
    /// * `data` - Array of string data to populate the chart cache.
    /// * `is_numeric` - The chart cache date is numeric.
    ///
    #[allow(dead_code)] // This is only used for internal testing.
    pub(crate) fn set_cache(
        &mut self,
        data: &[&str],
        cache_type: ChartRangeCacheDataType,
    ) -> &mut ChartRange {
        self.cache = ChartRangeCacheData {
            cache_type,
            data: data.iter().map(std::string::ToString::to_string).collect(),
        };
        self
    }
}

#[derive(Clone, PartialEq)]
pub(crate) struct ChartRangeCacheData {
    pub(crate) cache_type: ChartRangeCacheDataType,
    pub(crate) data: Vec<String>,
}

impl ChartRangeCacheData {
    pub(crate) fn new() -> ChartRangeCacheData {
        ChartRangeCacheData {
            cache_type: ChartRangeCacheDataType::None,
            data: vec![],
        }
    }

    pub(crate) fn has_data(&self) -> bool {
        !self.data.is_empty()
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum ChartRangeCacheDataType {
    None,
    String,
    Number,
    Date,
}

// -----------------------------------------------------------------------
// ChartType
// -----------------------------------------------------------------------

#[derive(Clone, Copy, PartialEq, Eq)]
/// The `ChartType` enum define the type of a [`Chart`] object.
///
/// The main original chart types are supported, see below.
///
/// Support for newer Excel chart types such as Treemap, Sunburst, Box and
/// Whisker, Statistical Histogram, Waterfall, Funnel and Maps is not currently
/// planned since the underlying structure is substantially different from the
/// implemented chart types.
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

    /// A Stock chart showing Open-High-Low-Close data. It is also possible to
    /// show High-Low-Close data.
    ///
    /// Note, Volume variants of the Excel stock charts aren't currently
    /// supported but will be in a future release.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_stock.png">
    Stock,
}

// -----------------------------------------------------------------------
// ChartTitle
// -----------------------------------------------------------------------

/// The `ChartTitle` struct represents a chart title.
///
/// It is used in conjunction with the [`Chart`] struct.
///
#[derive(Clone, PartialEq)]
pub struct ChartTitle {
    pub(crate) range: ChartRange,
    pub(crate) format: ChartFormat,
    pub(crate) font: ChartFont,
    name: String,
    hidden: bool,
    is_horizontal: bool,
    ignore_rich_para: bool,
}

impl ChartTitle {
    pub(crate) fn new() -> ChartTitle {
        ChartTitle {
            range: ChartRange::default(),
            format: ChartFormat::default(),
            font: ChartFont::default(),
            name: String::new(),
            hidden: false,
            is_horizontal: false,
            ignore_rich_para: false,
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
    /// # Parameters
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
    /// A simple chart example using the `rust_xlsxwriter` library.
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
    /// #
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
    /// - [`ChartFormat::set_solid_fill()`]: Set the [`ChartSolidFill`] properties.
    /// - [`ChartFormat::set_pattern_fill()`]: Set the [`ChartPatternFill`] properties.
    /// - [`ChartFormat::set_gradient_fill()`]: Set the [`ChartGradientFill`] properties.
    /// - [`ChartFormat::set_no_fill()`]: Turn off the fill for the chart object.
    /// - [`ChartFormat::set_line()`]: Set the [`ChartLine`] properties.
    /// - [`ChartFormat::set_border()`]: Set the [`ChartBorder`] properties.
    ///   A synonym for [`ChartLine`] depending on context.
    /// - [`ChartFormat::set_no_line()`]: Turn off the line for the chart object.
    /// - [`ChartFormat::set_no_border()`]: Turn off the border for the chart object.
    ///
    /// # Parameters
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

    /// Set the font properties of a chart title.
    ///
    /// Set the font properties of a chart title using a [`ChartFont`]
    /// reference. Example font properties that can be set are:
    ///
    /// - [`ChartFont::set_bold()`]
    /// - [`ChartFont::set_italic()`]
    /// - [`ChartFont::set_name()`]
    /// - [`ChartFont::set_size()`]
    /// - [`ChartFont::set_rotation()`]
    ///
    /// See [`ChartFont`] for full details.
    ///
    /// # Parameters
    ///
    /// `font`: A [`ChartFont`] struct reference to represent the font
    /// properties.
    ///
    /// # Examples
    ///
    /// An example of setting the font for a chart title.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_title_set_font.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFont, ChartType, Workbook, XlsxError};
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
    ///     // Add a data series.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font.
    ///     chart
    ///         .title()
    ///         .set_name("Title")
    ///         .set_font(ChartFont::new().set_bold().set_color("#FF0000"));
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_title_set_font.png">
    ///
    pub fn set_font(&mut self, font: &ChartFont) -> &mut ChartTitle {
        let mut font = font.clone();
        font.has_default_bold = true;

        if font.italic || font.is_latin() {
            font.has_baseline = true;
        }

        self.font = font;
        self
    }
}

// -----------------------------------------------------------------------
// ChartMarker
// -----------------------------------------------------------------------

/// The `ChartMarker` struct represents a chart marker.
///
/// The [`ChartMarker`] struct represents the properties of a marker on a Line,
/// Scatter or Radar chart. In Excel a marker is a shape that represents a data
/// point in a chart series.
///
/// It is used in conjunction with the [`Chart`] struct.
///
/// # Examples
///
/// An example of adding markers to a Line chart.
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
            format: ChartFormat::default(),
        }
    }

    /// Set the automatic/default marker type.
    ///
    /// Allow the marker type to be set automatically by Excel.
    ///
    /// # Examples
    ///
    /// An example of adding automatic markers to a Line chart.
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
    /// # Parameters
    ///
    /// `marker_type`: a [`ChartMarkerType`] enum value.
    ///
    /// # Examples
    ///
    /// An example of adding markers to a Line chart with user defined marker
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
    /// # Parameters
    ///
    /// `size` - The size of the marker.
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
    /// - [`ChartFormat::set_solid_fill()`]: Set the [`ChartSolidFill`] properties.
    /// - [`ChartFormat::set_pattern_fill()`]: Set the [`ChartPatternFill`] properties.
    /// - [`ChartFormat::set_gradient_fill()`]: Set the [`ChartGradientFill`] properties.
    /// - [`ChartFormat::set_no_fill()`]: Turn off the fill for the chart object.
    /// - [`ChartFormat::set_line()`]: Set the [`ChartLine`] properties.
    /// - [`ChartFormat::set_border()`]: Set the [`ChartBorder`] properties.
    ///   A synonym for [`ChartLine`] depending on context.
    /// - [`ChartFormat::set_no_line()`]: Turn off the line for the chart object.
    /// - [`ChartFormat::set_no_border()`]: Turn off the border for the chart object.
    /// - `set_no_border`: Turn off the border for the chart object.
    ///
    /// # Parameters
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

/// The `ChartMarkerType` enum defines the [`Chart`] marker types.
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

    /// Short dash marker type. This is also called `dot` in some Excel versions.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_short_dash.png">
    ShortDash,

    /// Long dash marker type. This is also called `dash` in some Excel versions.
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

impl fmt::Display for ChartMarkerType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::X => write!(f, "x"),
            Self::Star => write!(f, "star"),
            Self::Circle => write!(f, "circle"),
            Self::Square => write!(f, "square"),
            Self::Diamond => write!(f, "diamond"),
            Self::LongDash => write!(f, "dash"),
            Self::PlusSign => write!(f, "plus"),
            Self::Triangle => write!(f, "triangle"),
            Self::ShortDash => write!(f, "dot"),
        }
    }
}

// -----------------------------------------------------------------------
// ChartDataLabel
// -----------------------------------------------------------------------

/// The `ChartDataLabel` struct represents a chart data label.
///
/// The [`ChartDataLabel`] struct represents the properties of the data labels
/// for a chart series. In Excel a data label can be added to a chart series to
/// indicate the values of the plotted data points. It can also be used to
/// indicate other properties such as the category or, for Pie charts, the
/// percentage.
///
/// It is used in conjunction with the [`Chart`] struct.
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/chart_data_labels_dialog.png">
///
///
/// # Examples
///
/// An example of adding data labels to a chart series.
///
/// ```
/// # // This code is available in examples/doc_chart_data_labels.rs
/// #
/// use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError, ChartDataLabel};
///
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///     let worksheet = workbook.add_worksheet();
///
///     // Add some data for the chart.
///     worksheet.write(0, 0, 10)?;
///     worksheet.write(1, 0, 40)?;
///     worksheet.write(2, 0, 50)?;
///     worksheet.write(3, 0, 20)?;
///     worksheet.write(4, 0, 10)?;
///     worksheet.write(5, 0, 50)?;
///
///     // Create a new chart.
///     let mut chart = Chart::new(ChartType::Column);
///
///     // Add a data series.
///     chart
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_data_label(ChartDataLabel::new().show_value());
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
/// <img src="https://rustxlsxwriter.github.io/images/chart_data_labels.png">
///
#[derive(Clone, PartialEq)]
pub struct ChartDataLabel {
    pub(crate) format: ChartFormat,
    pub(crate) show_value: bool,
    pub(crate) show_category_name: bool,
    pub(crate) show_series_name: bool,
    pub(crate) show_leader_lines: bool,
    pub(crate) show_legend_key: bool,
    pub(crate) show_percentage: bool,
    pub(crate) position: ChartDataLabelPosition,
    pub(crate) separator: char,
    pub(crate) title: ChartTitle,
    pub(crate) is_hidden: bool,
    pub(crate) is_custom: bool,
    pub(crate) font: Option<ChartFont>,
    pub(crate) num_format: String,
}

impl Default for ChartDataLabel {
    fn default() -> Self {
        Self::new()
    }
}

impl ChartDataLabel {
    /// Create a new `ChartDataLabel` object to represent a Chart data label.
    ///
    pub fn new() -> ChartDataLabel {
        ChartDataLabel {
            format: ChartFormat::default(),
            show_value: false,
            show_category_name: false,
            show_series_name: false,
            show_leader_lines: false,
            show_legend_key: false,
            show_percentage: false,
            position: ChartDataLabelPosition::Default,
            separator: ',',
            title: ChartTitle::new(),
            is_hidden: false,
            is_custom: false,
            font: None,
            num_format: String::new(),
        }
    }

    /// Display the point value on the data label.
    ///
    /// This method turns on the option to display the value of the data point.
    ///
    /// If no other display options is set, such as `show_category()`, then this
    /// value will default to on, like in Excel.
    ///
    /// # Examples
    ///
    /// An example of adding data labels to a chart series.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_data_labels.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError, ChartDataLabel};
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
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_data_label(ChartDataLabel::new().show_value());
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_data_labels.png">
    ///
    pub fn show_value(&mut self) -> &mut ChartDataLabel {
        self.show_value = true;
        self
    }

    /// Display the point category name on the data label.
    ///
    /// This method turns on the option to display the category name of the data
    /// point.
    ///
    /// # Examples
    ///
    /// An example of adding data labels to a chart series with value and category
    /// details.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_data_labels_show_category_name.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartDataLabel, ChartType, Workbook, XlsxError};
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
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_data_label(ChartDataLabel::new().show_value().show_category_name());
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_data_labels_show_category_name.png">
    ///
    pub fn show_category_name(&mut self) -> &mut ChartDataLabel {
        self.show_category_name = true;
        self
    }

    /// Display the series name on the data label.
    ///
    pub fn show_series_name(&mut self) -> &mut ChartDataLabel {
        self.show_series_name = true;
        self
    }

    /// Display leader lines from/to the data label.
    ///
    /// **Note**: Even when leader lines are turned on they are not
    /// automatically visible in a chart. In Excel a leader line only appears if
    /// the data label is moved manually or if the data labels are very close
    /// and need to be adjusted automatically.
    ///
    pub fn show_leader_lines(&mut self) -> &mut ChartDataLabel {
        self.show_leader_lines = true;
        self
    }

    /// Show the legend key/symbol on the data label.
    ///
    pub fn show_legend_key(&mut self) -> &mut ChartDataLabel {
        self.show_legend_key = true;
        self
    }

    /// Display the chart value as a percentage.
    ///
    /// This method is used to turn on the display of data labels as a
    /// percentage for a series. It is mainly used for pie charts.
    ///
    /// # Examples
    ///
    /// An example of setting the percentage for the data labels of a chart
    /// series. Usually this only applies to a Pie or Doughnut chart.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_data_labels_show_percentage.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartDataLabel, Workbook, XlsxError};
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
    ///     let mut chart = Chart::new_pie();
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$3")
    ///         .set_data_label(ChartDataLabel::new().show_percentage());
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_data_labels_show_percentage.png">
    ///
    pub fn show_percentage(&mut self) -> &mut ChartDataLabel {
        self.show_percentage = true;
        self
    }

    /// Set the default position of the data label.
    ///
    /// In Excel the available data label positions vary for different chart
    /// types. The available, and default, positions are shown below with their
    /// [`ChartDataLabel`] value:
    ///
    /// | Position     | Line, Scatter | Bar, Column   | Pie, Doughnut | Area, Radar   |
    /// | :----------- | :------------ | :------------ | :------------ | :------------ |
    /// | `Center`     | Yes           | Yes           | Yes           | Yes (default) |
    /// | `Right`      | Yes (default) |               |               |               |
    /// | `Left`       | Yes           |               |               |               |
    /// | `Above`      | Yes           |               |               |               |
    /// | `Below`      | Yes           |               |               |               |
    /// | `InsideBase` |               | Yes           |               |               |
    /// | `InsideEnd`  |               | Yes           | Yes           |               |
    /// | `OutsideEnd` |               | Yes (default) | Yes           |               |
    /// | `BestFit`    |               |               | Yes (default) |               |
    ///
    /// # Parameters
    ///
    /// `position`: The label position as defined by the [`ChartDataLabel`] values.
    ///
    /// # Examples
    ///
    /// An example of adding data labels to a chart series and changing their
    /// default position.
    ///
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_data_labels_set_position.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartDataLabel, ChartDataLabelPosition, ChartType, Workbook, XlsxError,
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
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_data_label(
    ///             ChartDataLabel::new()
    ///                 .show_value()
    ///                 .set_position(ChartDataLabelPosition::InsideEnd),
    ///         );
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_data_labels_set_position.png">
    ///
    pub fn set_position(&mut self, position: ChartDataLabelPosition) -> &mut ChartDataLabel {
        self.position = position;
        self
    }

    /// Set the formatting properties for a chart data label.
    ///
    /// Set the formatting properties for a chart data label via a [`ChartFormat`]
    /// object or a sub struct that implements [`IntoChartFormat`].
    ///
    /// The formatting that can be applied via a [`ChartFormat`] object are:
    /// - [`ChartFormat::set_solid_fill()`]: Set the [`ChartSolidFill`] properties.
    /// - [`ChartFormat::set_pattern_fill()`]: Set the [`ChartPatternFill`] properties.
    /// - [`ChartFormat::set_gradient_fill()`]: Set the [`ChartGradientFill`] properties.
    /// - [`ChartFormat::set_no_fill()`]: Turn off the fill for the chart object.
    /// - [`ChartFormat::set_line()`]: Set the [`ChartLine`] properties.
    /// - [`ChartFormat::set_border()`]: Set the [`ChartBorder`] properties.
    ///   A synonym for [`ChartLine`] depending on context.
    /// - [`ChartFormat::set_no_line()`]: Turn off the line for the chart object.
    /// - [`ChartFormat::set_no_border()`]: Turn off the border for the chart object.
    /// - `set_no_border`: Turn off the border for the chart object.
    ///
    /// # Parameters
    ///
    /// `format`: A [`ChartFormat`] struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// [`IntoChartFormat`] for details.
    ///
    /// # Examples
    ///
    /// An example of adding data labels to a chart series with formatting.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_data_labels_set_format.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartDataLabel, ChartFormat, ChartLine, ChartSolidFill, ChartType, Workbook, XlsxError,
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
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_data_label(
    ///             ChartDataLabel::new().show_value().set_format(
    ///                 ChartFormat::new()
    ///                     .set_border(ChartLine::new().set_color("#FF0000"))
    ///                     .set_solid_fill(ChartSolidFill::new().set_color("#FFFF00")),
    ///             ),
    ///         );
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_data_labels_set_format.png">
    ///
    pub fn set_format<T>(&mut self, format: T) -> &mut ChartDataLabel
    where
        T: IntoChartFormat,
    {
        self.format = format.new_chart_format();
        self.title.ignore_rich_para = false;
        self
    }

    /// Set the font properties of a chart data label.
    ///
    /// Set the font properties of a chart data labels using a [`ChartFont`]
    /// reference. Example font properties that can be set are:
    ///
    /// - [`ChartFont::set_bold()`]
    /// - [`ChartFont::set_italic()`]
    /// - [`ChartFont::set_name()`]
    /// - [`ChartFont::set_size()`]
    /// - [`ChartFont::set_rotation()`]
    ///
    /// See [`ChartFont`] for full details. ///
    /// # Parameters
    ///
    /// `font`: A [`ChartFont`] struct reference to represent the font
    /// properties.
    ///
    ///
    /// # Examples
    ///
    /// An example of adding data labels to a chart series with font formatting.
    ///
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_data_labels_set_font.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartDataLabel, ChartFont, ChartType, Workbook, XlsxError};
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
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_data_label(
    ///             ChartDataLabel::new()
    ///                 .show_value()
    ///                 .set_font(ChartFont::new().set_bold().set_color("#FF0000")),
    ///         );
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_data_labels_set_font.png">
    ///
    pub fn set_font(&mut self, font: &ChartFont) -> &mut ChartDataLabel {
        let mut font = font.clone();

        if font.italic {
            font.has_baseline = true;
        }

        self.font = Some(font);
        self
    }

    /// Set the number format for a chart data label.
    ///
    /// Excel plots/displays data in charts with the same number formatting that
    /// it has in the worksheet. The `set_num_format()` method allows you to
    /// override this and controls whether a number is displayed as an integer,
    /// a floating point number, a date, a currency value or some other user
    /// defined format.
    ///
    /// See also [Number Format Categories] and [Number Formats in different
    /// locales] in the documentation for [`Format`](crate::Format).
    ///
    /// [Number Format Categories]: crate::Format#number-format-categories
    /// [Number Formats in different locales]:
    ///     crate::Format#number-formats-in-different-locales
    ///
    /// # Parameters
    ///
    /// * `num_format` - The number format property.
    ///
    ///
    /// # Examples
    ///
    /// An example of adding data labels to a chart series with number formatting.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_data_labels_set_num_format.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartDataLabel, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 0.1)?;
    /// #     worksheet.write(1, 0, 0.4)?;
    /// #     worksheet.write(2, 0, 0.5)?;
    /// #     worksheet.write(3, 0, 0.2)?;
    /// #     worksheet.write(4, 0, 0.1)?;
    /// #     worksheet.write(5, 0, 0.5)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_data_label(ChartDataLabel::new().show_value().set_num_format("0.00%"));
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_data_labels_set_num_format.png">
    ///
    pub fn set_num_format(&mut self, num_format: impl Into<String>) -> &mut ChartDataLabel {
        self.num_format = num_format.into();
        self
    }

    /// Set the separator for the displayed values of the data label.
    ///
    /// The allowable separators are `','` (comma), `';'` (semicolon), `'.'`
    /// (full stop), `'\n'` (new line) and `' '` (space).
    ///
    /// # Parameters
    ///
    /// `separator` - The label separator character.
    ///
    pub fn set_separator(&mut self, separator: char) -> &mut ChartDataLabel {
        // Accept valid separators only apart from comma which is the default.
        if ";. \n".contains(separator) {
            self.separator = separator;
        }

        self
    }

    /// Display the point Y value on the data label.
    ///
    /// This is the same as [`show_value()`](ChartDataLabel::show_value) except
    /// it is named differently in Excel for Scatter charts. The methods are
    /// equivalent
    /// and either one can be used for any chart type.
    ///
    pub fn show_y_value(&mut self) -> &mut ChartDataLabel {
        self.show_value()
    }

    /// Display the point X value on the data label.
    ///
    /// This is the same as
    /// [`show_category_name()`](ChartDataLabel::show_category_name) except it
    /// is named differently in Excel for Scatter charts. The methods are
    /// equivalent and either one can be used for any chart type.
    ///
    pub fn show_x_value(&mut self) -> &mut ChartDataLabel {
        self.show_category_name()
    }

    /// Set the value for a custom data label.
    ///
    /// This method sets the value of a custom data label used with the
    /// [`set_custom_data_labels()`](ChartSeries::set_custom_data_labels)
    /// method. It is ignored if used with a series [`ChartDataLabel`].
    ///
    /// # Parameters
    ///
    /// * `value` - A [`IntoChartRange`] property which can be one of the
    ///   following generic types:
    ///    - A simple string title.
    ///    - A string with an Excel like range formula such as `"Sheet1!$A$1"`.
    ///    - A tuple that can be used to create the range programmatically using
    ///      a sheet name and zero indexed row and column values like:
    ///      `("Sheet1", 0, 0)` (this gives the same range as the previous
    ///      string value).
    ///
    /// # Examples
    ///
    /// An example of adding custom data labels to a chart series. This is
    /// useful when you want to label the points of a data series with
    /// information that isn't contained in the value or category names.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_custom_data_labels1.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartDataLabel, ChartType, Workbook, XlsxError};
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
    ///     // Create some custom data labels.
    ///     let data_labels = [
    ///         ChartDataLabel::new().set_value("Alice").to_custom(),
    ///         ChartDataLabel::new().set_value("Bob").to_custom(),
    ///         ChartDataLabel::new().set_value("Carol").to_custom(),
    ///         ChartDataLabel::new().set_value("Dave").to_custom(),
    ///         ChartDataLabel::new().set_value("Eve").to_custom(),
    ///         ChartDataLabel::new().set_value("Frank").to_custom(),
    ///     ];
    ///
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_custom_data_labels(&data_labels);
    ///
    ///     // Turn legend off for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_set_custom_data_labels1.png">
    ///
    ///
    /// This example shows how to get the data from cells. In Excel this is a
    /// single command called "Value from Cells" but in `rust_xlsxwriter` it
    /// needs to be broken down into a cell reference for each data label.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_custom_data_labels2.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartDataLabel, ChartType, Workbook, XlsxError};
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
    /// #     worksheet.write(0, 1, "Asia")?;
    /// #     worksheet.write(1, 1, "Africa")?;
    /// #     worksheet.write(2, 1, "Europe")?;
    /// #     worksheet.write(3, 1, "Americas")?;
    /// #     worksheet.write(4, 1, "Oceania")?;
    /// #     worksheet.write(5, 1, "Antarctic")?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Create some custom data labels.
    ///     let data_labels = [
    ///         ChartDataLabel::new().set_value("=Sheet1!$B$1").to_custom(),
    ///         ChartDataLabel::new().set_value("=Sheet1!$B$2").to_custom(),
    ///         ChartDataLabel::new().set_value("=Sheet1!$B$3").to_custom(),
    ///         ChartDataLabel::new().set_value("=Sheet1!$B$4").to_custom(),
    ///         ChartDataLabel::new().set_value("=Sheet1!$B$5").to_custom(),
    ///         ChartDataLabel::new().set_value("=Sheet1!$B$6").to_custom(),
    ///     ];
    ///
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_custom_data_labels(&data_labels);
    ///
    ///     // Turn legend off for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_set_custom_data_labels2.png">
    ///
    ///
    pub fn set_value<T>(&mut self, value: T) -> &mut ChartDataLabel
    where
        T: IntoChartRange,
    {
        self.title.set_name(value);
        self.title.ignore_rich_para = true;
        self.show_value = true;
        self
    }

    /// Set a custom data label as hidden.
    ///
    /// This method hides a custom data label used with the
    /// [`set_custom_data_labels()`](ChartSeries::set_custom_data_labels)
    /// method. It is ignored if used with a series [`ChartDataLabel`].
    ///
    /// # Examples
    ///
    /// An example of adding custom data labels to a chart series.
    ///
    /// This example shows how to add default/non-custom data labels along with
    /// custom data labels. This is done in two ways: with an explicit
    /// `default()` data label and with an implicit default for points that
    /// aren't covered at the end of the list.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_set_custom_data_labels3.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartDataLabel, ChartType, Workbook, XlsxError};
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
    ///     // Create some custom data labels.
    ///     let data_labels = [
    ///         ChartDataLabel::default(),
    ///         ChartDataLabel::default(),
    ///         ChartDataLabel::new().set_value("Alice").to_custom(),
    ///         ChartDataLabel::new().set_value("Bob").to_custom(),
    ///         // All other points after this will get a default label.
    ///     ];
    ///
    ///     // Add a data series.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_custom_data_labels(&data_labels);
    ///
    ///     // Turn legend off for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_set_custom_data_labels3.png">
    ///
    pub fn set_hidden(&mut self) -> &mut ChartDataLabel {
        self.is_hidden = true;
        self
    }

    /// Turn a data label reference into a custom data label.
    ///
    /// Converts a `&ChartDataLabel` reference into a [`ChartDataLabel`] so that
    /// it can be used as a custom data label with the
    /// [`set_custom_data_labels()`](ChartSeries::set_custom_data_labels)
    /// method.
    ///
    /// This is a syntactic shortcut for a simple `clone()`.
    ///
    pub fn to_custom(&mut self) -> ChartDataLabel {
        self.clone()
    }

    // Check if the data label is in the default/unmodified condition.
    pub(crate) fn is_default(&self) -> bool {
        lazy_static! {
            static ref DEFAULT_STATE: ChartDataLabel = ChartDataLabel::default();
        };
        self == &*DEFAULT_STATE
    }
}

/// The `ChartMarkerType` enum defines the [`Chart`] data label positions.
///
/// In Excel the available data label positions vary for different chart
/// types. The available, and default, positions are:
///
/// | Position      | Line, Scatter | Bar, Column   | Pie, Doughnut | Area, Radar   |
/// | :------------ | :------------ | :------------ | :------------ | :------------ |
/// | `Center`      | Yes           | Yes           | Yes           | Yes (default) |
/// | `Right`       | Yes (default) |               |               |               |
/// | `Left`        | Yes           |               |               |               |
/// | `Above`       | Yes           |               |               |               |
/// | `Below`       | Yes           |               |               |               |
/// | `InsideBase`  |               | Yes           |               |               |
/// | `InsideEnd`   |               | Yes           | Yes           |               |
/// | `OutsideEnd`  |               | Yes (default) | Yes           |               |
/// | `BestFit`     |               |               | Yes (default) |               |
///
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ChartDataLabelPosition {
    /// Series data label position: Default position.
    Default,

    /// Series data label position: Center.
    Center,

    /// Series data label position: Right.
    Right,

    /// Series data label position: Left.
    Left,

    /// Series data label position: Above.
    Above,

    /// Series data label position: Below.
    Below,

    /// Series data label position: Inside base.
    InsideBase,

    /// Series data label position: Inside end.
    InsideEnd,

    /// Series data label position: Outside end.
    OutsideEnd,

    /// Series data label position: Best fit.
    BestFit,
}

impl fmt::Display for ChartDataLabelPosition {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Left => write!(f, "l"),
            Self::Right => write!(f, "r"),
            Self::Above => write!(f, "t"),
            Self::Below => write!(f, "b"),
            Self::Center => write!(f, "ctr"),
            Self::Default => write!(f, ""),
            Self::BestFit => write!(f, "bestFit"),
            Self::InsideEnd => write!(f, "inEnd"),
            Self::InsideBase => write!(f, "inBase"),
            Self::OutsideEnd => write!(f, "outEnd"),
        }
    }
}

// -----------------------------------------------------------------------
// ChartPoint
// -----------------------------------------------------------------------

/// The `ChartPoint` struct represents a chart point.
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
/// It is used in conjunction with the [`Chart`] struct.
///
/// # Examples
///
/// An example of formatting the individual segments of a Pie chart.
///
/// ```
/// # // This code is available in examples/doc_chart_set_points.rs
/// #
/// # use rust_xlsxwriter::{
/// #     Chart, ChartFormat, ChartPoint, ChartSolidFill, Workbook, XlsxError,
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
///     let mut chart = Chart::new_pie();
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
            format: ChartFormat::default(),
        }
    }

    /// Set the formatting properties for a chart point.
    ///
    /// Set the formatting properties for a chart point via a [`ChartFormat`]
    /// object or a sub struct that implements [`IntoChartFormat`].
    ///
    /// The formatting that can be applied via a [`ChartFormat`] object are:
    ///
    /// - [`ChartFormat::set_solid_fill()`]: Set the [`ChartSolidFill`] properties.
    /// - [`ChartFormat::set_pattern_fill()`]: Set the [`ChartPatternFill`] properties.
    /// - [`ChartFormat::set_gradient_fill()`]: Set the [`ChartGradientFill`] properties.
    /// - [`ChartFormat::set_no_fill()`]: Turn off the fill for the chart object.
    /// - [`ChartFormat::set_line()`]: Set the [`ChartLine`] properties.
    /// - [`ChartFormat::set_border()`]: Set the [`ChartBorder`] properties.
    ///   A synonym for [`ChartLine`] depending on context.
    /// - [`ChartFormat::set_no_line()`]: Turn off the line for the chart object.
    /// - [`ChartFormat::set_no_border()`]: Turn off the border for the chart object.
    ///
    /// # Parameters
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

// -----------------------------------------------------------------------
// ChartAxis
// -----------------------------------------------------------------------

/// The `ChartAxis` struct represents a chart axis.
///
/// Used in conjunction with the [`Chart::x_axis()`] and [`Chart::y_axis()`].
///
/// It is used in conjunction with the [`Chart`] struct.
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
///
#[derive(Clone)]
pub struct ChartAxis {
    axis_type: ChartAxisType,
    axis_position: ChartAxisPosition,
    label_position: ChartAxisLabelPosition,
    pub(crate) title: ChartTitle,
    pub(crate) format: ChartFormat,
    pub(crate) font: Option<ChartFont>,
    pub(crate) num_format: String,
    pub(crate) num_format_linked_to_source: bool,
    pub(crate) reverse: bool,
    pub(crate) is_hidden: bool,
    pub(crate) automatic: bool,
    pub(crate) position_between_ticks: bool,
    pub(crate) max: String,
    pub(crate) min: String,
    pub(crate) major_unit: String,
    pub(crate) minor_unit: String,
    pub(crate) major_gridlines: bool,
    pub(crate) minor_gridlines: bool,
    pub(crate) major_gridlines_line: Option<ChartLine>,
    pub(crate) minor_gridlines_line: Option<ChartLine>,
    pub(crate) log_base: u16,
    pub(crate) label_interval: u16,
    pub(crate) tick_interval: u16,
    pub(crate) major_tick_type: Option<ChartAxisTickType>,
    pub(crate) minor_tick_type: Option<ChartAxisTickType>,
    pub(crate) major_unit_date_type: Option<ChartAxisDateUnitType>,
    pub(crate) minor_unit_date_type: Option<ChartAxisDateUnitType>,
    pub(crate) display_units_type: ChartAxisDisplayUnitType,
    pub(crate) display_units_visible: bool,
    pub(crate) crossing: ChartAxisCrossing,
    pub(crate) label_alignment: ChartAxisLabelAlignment,
}

impl ChartAxis {
    pub(crate) fn new() -> ChartAxis {
        ChartAxis {
            axis_type: ChartAxisType::Value,
            axis_position: ChartAxisPosition::Bottom,
            label_position: ChartAxisLabelPosition::NextTo,
            title: ChartTitle::new(),
            format: ChartFormat::default(),
            font: None,
            num_format: String::new(),
            num_format_linked_to_source: false,
            reverse: false,
            is_hidden: false,
            automatic: false,
            position_between_ticks: true,
            max: String::new(),
            min: String::new(),
            major_unit: String::new(),
            minor_unit: String::new(),
            major_gridlines: false,
            minor_gridlines: false,
            major_gridlines_line: None,
            minor_gridlines_line: None,
            log_base: 0,
            label_interval: 0,
            tick_interval: 0,
            major_tick_type: None,
            minor_tick_type: None,
            major_unit_date_type: None,
            minor_unit_date_type: None,
            display_units_type: ChartAxisDisplayUnitType::None,
            display_units_visible: false,
            crossing: ChartAxisCrossing::Automatic,
            label_alignment: ChartAxisLabelAlignment::Center,
        }
    }

    /// Add a title for a chart axis.
    ///
    /// Set the name (title) for the chart axis.
    ///
    /// The name can be a simple string, a formula such as `Sheet1!$A$1` or a
    /// tuple with a sheet name, row and column such as `('Sheet1', 0, 0)`.
    ///
    /// # Parameters
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_name.png">
    ///
    pub fn set_name<T>(&mut self, name: T) -> &mut ChartAxis
    where
        T: IntoChartRange,
    {
        self.title.set_name(name);
        self
    }

    /// Set the font properties of a chart axis title.
    ///
    /// Set the font properties of a chart axis name/title using a [`ChartFont`]
    /// reference. Example font properties that can be set are:
    ///
    /// - [`ChartFont::set_bold()`]
    /// - [`ChartFont::set_italic()`]
    /// - [`ChartFont::set_name()`]
    /// - [`ChartFont::set_size()`]
    /// - [`ChartFont::set_rotation()`]
    ///
    /// See [`ChartFont`] for full details.
    ///
    ///
    /// The name font property for an axis represents the font for
    /// the axis title. To set the font for the category or value numbers use
    /// the [`set_font()`](ChartAxis::set_font) method.
    ///
    /// # Parameters
    ///
    /// `font`: A [`ChartFont`] struct reference to represent the font
    /// properties.
    ///
    ///
    /// # Examples
    ///
    /// An example of setting the font for a chart axis title.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_name_font.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFont, ChartType, Workbook, XlsxError};
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
    ///     // Add a data series.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font.
    ///     chart
    ///         .x_axis()
    ///         .set_name("X-Axis")
    ///         .set_name_font(ChartFont::new().set_bold().set_color("#FF0000"));
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_name_font.png">
    ///
    pub fn set_name_font(&mut self, font: &ChartFont) -> &mut ChartAxis {
        self.title.set_font(font);
        self
    }

    /// Set the formatting properties for a chart axis.
    ///
    /// Set the formatting properties for a chart axis via a [`ChartFormat`]
    /// object or a sub struct that implements [`IntoChartFormat`].
    ///
    /// The formatting that can be applied via a [`ChartFormat`] object are:
    ///
    /// - [`ChartFormat::set_solid_fill()`]: Set the [`ChartSolidFill`] properties.
    /// - [`ChartFormat::set_pattern_fill()`]: Set the [`ChartPatternFill`] properties.
    /// - [`ChartFormat::set_gradient_fill()`]: Set the [`ChartGradientFill`] properties.
    /// - [`ChartFormat::set_no_fill()`]: Turn off the fill for the chart object.
    /// - [`ChartFormat::set_line()`]: Set the [`ChartLine`] properties.
    /// - [`ChartFormat::set_border()`]: Set the [`ChartBorder`] properties.
    ///   A synonym for [`ChartLine`] depending on context.
    /// - [`ChartFormat::set_no_line()`]: Turn off the line for the chart object.
    /// - [`ChartFormat::set_no_border()`]: Turn off the border for the chart object.
    ///
    /// # Parameters
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

    /// Set the font properties of a chart axis.
    ///
    /// Set the font properties of a chart axis using a [`ChartFont`] reference.
    /// Example font properties that can be set are:
    ///
    /// - [`ChartFont::set_bold()`]
    /// - [`ChartFont::set_italic()`]
    /// - [`ChartFont::set_name()`]
    /// - [`ChartFont::set_size()`]
    /// - [`ChartFont::set_rotation()`]
    ///
    /// See [`ChartFont`] for full details.
    ///
    /// The font property for an axis represents the font for the category or
    /// value names or numbers. To set the font for the axis name/title use the
    /// [`set_name_font()`](ChartAxis::set_name_font) method.
    ///
    /// # Parameters
    ///
    /// `font`: A [`ChartFont`] struct reference to represent the font
    /// properties.
    ///
    /// # Examples
    ///
    /// An example of setting the font for a chart axis.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_font.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFont, ChartType, Workbook, XlsxError};
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
    ///     // Add a data series.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font for an axis.
    ///     chart.x_axis().set_font(
    ///         ChartFont::new()
    ///             .set_bold()
    ///             .set_italic()
    ///             .set_name("Consolas")
    ///             .set_color("#FF0000"),
    ///     );
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_font.png">
    ///
    pub fn set_font(&mut self, font: &ChartFont) -> &mut ChartAxis {
        let mut font = font.clone();

        if font.italic || font.is_latin() {
            font.has_baseline = true;
        }

        if font.italic && font.bold.is_none() {
            font.bold = Some(false);
        }

        self.font = Some(font);
        self
    }

    /// Set the number format for a chart axis.
    ///
    /// Excel plots/displays data in charts with the same number formatting that
    /// it has in the worksheet. The `set_num_format()` method allows you to
    /// override this and controls whether a number is displayed as an integer,
    /// a floating point number, a date, a currency value or some other user
    /// defined format.
    ///
    /// See also [Number Format Categories] and [Number Formats in different
    /// locales] in the documentation for [`Format`](crate::Format).
    ///
    /// [Number Format Categories]: crate::Format#number-format-categories
    /// [Number Formats in different locales]:
    ///     crate::Format#number-formats-in-different-locales
    ///
    /// # Parameters
    ///
    /// * `num_format` - The number format property.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the number format a chart axes.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_num_format.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 0.1)?;
    /// #     worksheet.write(1, 0, 0.4)?;
    /// #     worksheet.write(2, 0, 0.5)?;
    /// #     worksheet.write(3, 0, 0.2)?;
    /// #     worksheet.write(4, 0, 0.1)?;
    /// #     worksheet.write(5, 0, 0.5)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the chart axis number format.
    ///     chart.y_axis().set_num_format("0.00%");
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_axis_set_num_format.png">
    ///
    pub fn set_num_format(&mut self, num_format: impl Into<String>) -> &mut ChartAxis {
        self.num_format = num_format.into();
        self
    }

    /// Set the category axis as a date axis.
    ///
    /// In general the "Category" axis (usually the X-axis) in Excel charts is
    /// made up of evenly spaced categories. This type of axis doesn't support
    /// features such as maximum and minimum even if the categories are numbers.
    /// The two exceptions to this are the "Value" axes used in Scatter charts
    /// and "Date" axes. Date axes are a combination of "Category" and "Value"
    /// axes and they support features of both types of axes.
    ///
    /// In order to have a date axes in your chart you need to have a range of
    /// Date/Time values in a worksheet that the
    /// [`ChartSeries::set_categories()`] refer to. You can then use the
    /// `set_date_axis()` method turns on the "date axis" property for a chart
    /// axis.
    ///
    /// See [Chart Value and Category Axes] for an explanation of the
    /// difference between Value and Category axes in Excel.
    ///
    /// [Chart Value and Category Axes]:
    ///     crate::chart#chart-value-and-category-axes
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting a date axis for a chart.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_date_axis.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #     let date_format = Format::new().set_num_format("yyyy-mm-dd");
    /// #
    /// #     // Adjust the date column width for clarity.
    /// #     worksheet.set_column_width(0, 11)?;
    /// #
    /// #     // Add some data for the chart.
    /// #     let dates = [
    /// #         ExcelDateTime::parse_from_str("2024-01-01")?,
    /// #         ExcelDateTime::parse_from_str("2024-01-02")?,
    /// #         ExcelDateTime::parse_from_str("2024-01-03")?,
    /// #         ExcelDateTime::parse_from_str("2024-01-04")?,
    /// #         ExcelDateTime::parse_from_str("2024-01-05")?,
    /// #     ];
    /// #     let values = [27.2, 25.03, 19.05, 20.34, 18.5];
    /// #
    /// #     worksheet.write_column_with_format(0, 0, dates, &date_format)?;
    /// #     worksheet.write_column(0, 1, values)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 0, 4, 0))
    ///         .set_values(("Sheet1", 0, 1, 4, 1));
    ///
    ///     // Set the axis as a date axis.
    ///     chart.x_axis().set_date_axis(true);
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
    /// src="https://rustxlsxwriter.github.io/images/chart_axis_set_date_axis.png">
    ///
    ///
    pub fn set_date_axis(&mut self, enable: bool) -> &mut ChartAxis {
        if enable {
            self.axis_type = ChartAxisType::Date;
        } else {
            self.axis_type = ChartAxisType::Category;
        }

        self.automatic = !enable;

        self
    }

    /// Set the category axis as a text axis.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn set_text_axis(&mut self, enable: bool) -> &mut ChartAxis {
        self.set_automatic_axis(enable);
        self
    }

    /// Set the category axis as an automatic axis - generally the default.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn set_automatic_axis(&mut self, enable: bool) -> &mut ChartAxis {
        self.automatic = enable;
        self
    }

    /// Set the crossing point for the opposite axis.
    ///
    /// By default Excel sets chart axes to cross at 0. If required you can use
    /// [`ChartAxis::set_crossing()`] and [`ChartAxisCrossing`] to define
    /// another point where the opposite axis will cross the current axis.
    ///
    /// The [`ChartAxisCrossing`] enum defines values like `max` and `min` but
    /// also allows you to define a category value for X-axes (except for
    /// Scatter and Date axes) and an actual value for Y-axes and Scatter and
    /// Date axes.
    ///
    /// # Parameters
    ///
    /// * `crossing` - A [`ChartAxisCrossing`] enum value.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the point where the axes will
    /// cross.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_crossing.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartAxisCrossing, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, "North")?;
    /// #     worksheet.write(1, 0, "South")?;
    /// #     worksheet.write(2, 0, "East")?;
    /// #     worksheet.write(3, 0, "West")?;
    /// #     worksheet.write(0, 1, 10)?;
    /// #     worksheet.write(1, 1, 35)?;
    /// #     worksheet.write(2, 1, 40)?;
    /// #     worksheet.write(3, 1, 25)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart
    ///         .add_series()
    ///         .set_categories("Sheet1!$A$1:$A$5")
    ///         .set_values("Sheet1!$B$1:$B$5");
    ///
    ///     // Set the X-axis crossing at a category index.
    ///     chart
    ///         .x_axis()
    ///         .set_crossing(ChartAxisCrossing::CategoryNumber(3));
    ///
    ///     // Set the Y-axis crossing at a value.
    ///     chart
    ///         .y_axis()
    ///         .set_crossing(ChartAxisCrossing::AxisValue(20.0));
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
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
    /// src="https://rustxlsxwriter.github.io/images/chart_axis_set_crossing1.png">
    ///
    /// For reference here is the default chart without default crossings:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_axis_set_crossing2.png">
    ///
    pub fn set_crossing(&mut self, crossing: ChartAxisCrossing) -> &mut ChartAxis {
        self.crossing = crossing;

        self
    }

    /// Reverse the direction of the axis categories or values.
    ///
    /// Reverse the direction that the axis data is plotted in from left to
    /// right or top to bottom.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating reversing the plotting direction of the
    /// chart axes.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_reverse.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(3, 0, 30)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Reverse the axis.
    ///     chart.x_axis().set_reverse();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_axis_set_reverse.png">
    ///
    pub fn set_reverse(&mut self) -> &mut ChartAxis {
        self.reverse = true;
        self
    }

    /// Set the maximum value for an axis.
    ///
    /// Set the maximum bound to be displayed for an axis.
    ///
    /// Maximum and minimum chart axis values can only be set for chart "Value"
    /// axes and "Category Date" axes in Excel. You cannot set a maximum or
    /// minimum value for "Category" axes even if the category values are
    /// numbers. See [Chart Value and Category Axes] for an explanation of the
    /// difference between Value and Category axes in Excel.
    ///
    /// See also [`ChartAxis::set_max_date()`] below.
    ///
    /// [Chart Value and Category Axes]:
    ///     crate::chart#chart-value-and-category-axes
    ///
    /// # Parameters
    ///
    /// `max` - The maximum bound for the axes.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the axes bounds for chart axes.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_max.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, -30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(3, 0, -30)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Set the value axes bounds.
    ///     chart.y_axis().set_min(-60);
    ///     chart.y_axis().set_max(60);
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_axis_set_max.png">
    ///
    pub fn set_max<T>(&mut self, max: T) -> &mut ChartAxis
    where
        T: Into<f64>,
    {
        self.max = max.into().to_string();
        self
    }

    /// Set the minimum value for an axis.
    ///
    /// Set the minimum bound to be displayed for an axis.
    ///
    /// See [`ChartAxis::set_max()`] above for a full explanation and example.
    ///
    /// # Parameters
    ///
    /// `min` - The minimum bound for the axes.
    ///
    pub fn set_min<T>(&mut self, min: T) -> &mut ChartAxis
    where
        T: Into<f64>,
    {
        self.min = min.into().to_string();
        self
    }

    /// Set the maximum date value for a date axis.
    ///
    /// Set the maximum date/time bound to be displayed for a date axis. This is
    /// just a syntactic helper around [`ChartAxis::set_max()`] to allow dates
    /// that support the [`IntoExcelDateTime`] trait to be passed to the API.
    ///
    /// # Parameters
    ///
    /// `datetime` - A date/time instance that implements [`IntoExcelDateTime`].
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the maximum and minimum values for a
    /// date axis.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_max_date.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #     let date_format = Format::new().set_num_format("yyyy-mm-dd");
    /// #
    /// #     // Adjust the date column width for clarity.
    /// #     worksheet.set_column_width(0, 11)?;
    /// #
    /// #     // Add some data for the chart.
    /// #     let dates = [
    /// #         ExcelDateTime::parse_from_str("2024-01-01")?,
    /// #         ExcelDateTime::parse_from_str("2024-01-02")?,
    /// #         ExcelDateTime::parse_from_str("2024-01-03")?,
    /// #         ExcelDateTime::parse_from_str("2024-01-04")?,
    /// #         ExcelDateTime::parse_from_str("2024-01-05")?,
    /// #         ExcelDateTime::parse_from_str("2024-01-06")?,
    /// #         ExcelDateTime::parse_from_str("2024-01-07")?,
    /// #     ];
    /// #     let values = [27, 25, 19, 20, 18, 15, 19];
    /// #
    /// #     worksheet.write_column_with_format(0, 0, dates, &date_format)?;
    /// #     worksheet.write_column(0, 1, values)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     chart
    ///         .add_series()
    ///         .set_categories(("Sheet1", 0, 0, 6, 0))
    ///         .set_values(("Sheet1", 0, 1, 6, 1));
    ///
    ///     // Set the axis as a date axis.
    ///     chart.x_axis().set_date_axis(true);
    ///
    ///     // Set the min and max date values for the chart.
    ///     let min_date = ExcelDateTime::parse_from_str("2024-01-02")?;
    ///     let max_date = ExcelDateTime::parse_from_str("2024-01-06")?;
    ///
    ///     chart.x_axis().set_min_date(min_date);
    ///     chart.x_axis().set_max_date(max_date);
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
    /// src="https://rustxlsxwriter.github.io/images/chart_axis_set_max_date.png">
    ///
    pub fn set_max_date(&mut self, datetime: impl IntoExcelDateTime) -> &mut ChartAxis {
        self.max = datetime.to_excel_serial_date().to_string();
        self
    }

    /// Set the minimum date value for a date axis.
    ///
    /// Set the minimum date/time bound to be displayed for a date axis. This is
    /// just a syntactic helper around [`ChartAxis::set_min()`] to allow dates
    /// that support the [`IntoExcelDateTime`] trait to be passed to the API.
    ///
    /// # Parameters
    ///
    /// `datetime` - A date/time instance that implements [`IntoExcelDateTime`].
    ///
    pub fn set_min_date(&mut self, datetime: impl IntoExcelDateTime) -> &mut ChartAxis {
        self.min = datetime.to_excel_serial_date().to_string();
        self
    }

    /// Set the increment of the major units in the axis range.
    ///
    /// Note, Excel only supports major/minor units for "Value" axes. In general
    /// you cannot set major/minor units for a X/Category axis even if the
    /// category values are numbers. See [Chart Value and Category Axes] for an
    /// explanation of the difference between Value and Category axes in Excel.
    ///
    /// [Chart Value and Category Axes]:
    ///     crate::chart#chart-value-and-category-axes
    ///
    /// # Parameters
    ///
    /// `value` - The major unit for the axes.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the units for chart axes.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_major_unit.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(3, 0, 30)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Turn on the minor gridlines.
    ///     chart.y_axis().set_minor_gridlines(true);
    ///
    ///     // Set the value axes major and minor units.
    ///     chart.y_axis().set_major_unit(20);
    ///     chart.y_axis().set_minor_unit(5);
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_major_unit.png">
    ///
    pub fn set_major_unit<T>(&mut self, value: T) -> &mut ChartAxis
    where
        T: Into<f64>,
    {
        let value = value.into();
        if value < 0.0 {
            eprintln!("Chart axis major unit '{value}' must be >= 0.0 in Excel");
            return self;
        }

        self.major_unit = value.to_string();
        self
    }

    /// Set the increment of the minor units in the axis range.
    ///
    /// See [`ChartAxis::set_major_unit()`] above for a full explanation and
    /// example.
    ///
    /// # Parameters
    ///
    /// `value` - The major unit for the axes.
    ///
    pub fn set_minor_unit<T>(&mut self, value: T) -> &mut ChartAxis
    where
        T: Into<f64>,
    {
        let value = value.into();
        if value < 0.0 {
            eprintln!("Chart axis minor unit '{value}' must be >= 0.0 in Excel");
            return self;
        }

        self.minor_unit = value.to_string();
        self
    }

    /// Set the display unit type such as Thousands, Millions, or other units.
    ///
    /// If the Value axis in your chart has very large numbers you can set the
    /// unit type to one of the following Excel values:
    ///
    /// - Hundreds
    /// - Thousands
    /// - Ten Thousands
    /// - Hundred Thousands
    /// - Millions
    /// - Ten Millions
    /// - Hundred Millions
    /// - Billions
    /// - Trillions
    ///
    /// # Parameters
    ///
    /// * `unit` - A [`ChartAxisDateUnitType`] enum value.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the units of the Value/Y-axis.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_display_unit_type.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartAxisDisplayUnitType, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 6_000_000)?;
    /// #     worksheet.write(1, 0, 17_000_000)?;
    /// #     worksheet.write(2, 0, 23_000_000)?;
    /// #     worksheet.write(3, 0, 4_000_000)?;
    /// #     worksheet.write(4, 0, 12_000_000)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Set the units for the axis.
    ///     chart
    ///         .y_axis()
    ///         .set_display_unit_type(ChartAxisDisplayUnitType::Millions);
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_display_unit_type.png">
    ///
    pub fn set_display_unit_type(&mut self, unit_type: ChartAxisDisplayUnitType) -> &mut ChartAxis {
        self.display_units_type = unit_type;
        self.display_units_visible = true;
        self
    }

    /// Make the display units visible (if they have been set).
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off.
    ///
    pub fn set_display_units_visible(&mut self, enable: bool) -> &mut ChartAxis {
        self.display_units_visible = enable;
        self
    }

    /// Set the major unit type as days, months or years.
    ///
    /// # Parameters
    ///
    /// * `unit` - A [`ChartAxisDateUnitType`] enum value.
    ///
    pub fn set_major_unit_date_type(&mut self, unit_type: ChartAxisDateUnitType) -> &mut ChartAxis {
        self.major_unit_date_type = Some(unit_type);
        self
    }

    /// Set the minor unit type as days, months or years.
    ///
    /// # Parameters
    ///
    /// * `unit` - A [`ChartAxisDateUnitType`] enum value.
    ///
    pub fn set_minor_unit_date_type(&mut self, unit_type: ChartAxisDateUnitType) -> &mut ChartAxis {
        self.minor_unit_date_type = Some(unit_type);
        self
    }

    /// Set the alignment of the axis labels relative to the tick mark.
    ///
    /// # Parameters
    ///
    /// * `unit` - A [`ChartAxisDateUnitType`] enum value.
    ///
    pub fn set_label_alignment(&mut self, alignment: ChartAxisLabelAlignment) -> &mut ChartAxis {
        self.label_alignment = alignment;
        self
    }

    /// Turn on/off major gridlines for a chart axis.
    ///
    /// Major gridlines are on by default for Y/Value axes but off for
    /// X/Category axes.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default for X axes.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating turning off the major gridlines for chart
    /// axes.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_major_gridlines.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(3, 0, 30)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Turn off the major gridlines, they are on by default.
    ///     chart.y_axis().set_major_gridlines(false);
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_axis_set_major_gridlines.png">
    ///
    pub fn set_major_gridlines(&mut self, enable: bool) -> &mut ChartAxis {
        self.major_gridlines = enable;
        self
    }

    /// Turn on/off minor gridlines for a chart axis.
    ///
    /// Minor gridlines are off by default.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating turning on the minor gridlines for chart axes.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_minor_gridlines.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(3, 0, 30)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Turn on the minor gridlines. The Y-axis major gridlines are on by
    ///     // default.
    ///     chart.y_axis().set_minor_gridlines(true);
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_minor_gridlines.png">
    ///
    pub fn set_minor_gridlines(&mut self, enable: bool) -> &mut ChartAxis {
        self.minor_gridlines = enable;
        self
    }

    /// Set the line formatting for a chart axis major gridlines.
    ///
    /// See the [`ChartLine`] struct for details on the line properties that can
    /// be set.
    ///
    /// # Parameters
    ///
    /// * `line` - A [`ChartLine`] struct reference.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating formatting the major gridlines for chart axes.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_major_gridlines_line.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartLine, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(3, 0, 30)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Format the major gridlines.
    ///     chart
    ///         .y_axis()
    ///         .set_major_gridlines_line(ChartLine::new().set_color("#FF0000"));
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_major_gridlines_line.png">
    ///
    pub fn set_major_gridlines_line(&mut self, line: &ChartLine) -> &mut ChartAxis {
        self.major_gridlines_line = Some(line.clone());
        self.major_gridlines = true;
        self
    }

    /// Set the line formatting for a chart axis minor gridlines.
    ///
    /// See the [`ChartLine`] struct for details on the line properties that can
    /// be set.
    ///
    /// # Parameters
    ///
    /// * `line` - A [`ChartLine`] struct reference.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating formatting the minor gridlines for chart axes.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_minor_gridlines_line.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartLine, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(3, 0, 30)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Format the minor gridlines.
    ///     chart
    ///         .y_axis()
    ///         .set_minor_gridlines_line(ChartLine::new().set_color("#FF0000"));
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_minor_gridlines_line.png">
    ///
    pub fn set_minor_gridlines_line(&mut self, line: &ChartLine) -> &mut ChartAxis {
        self.minor_gridlines_line = Some(line.clone());
        self.minor_gridlines = true;
        self
    }

    /// Set the label position for the axis.
    ///
    /// The label position defines where the values/categories for the axis are
    /// displayed. The position is controlled via [`ChartAxisLabelPosition`] enum.
    ///
    /// # Parameters
    ///
    /// * `position` - A [`ChartAxisLabelPosition`] enum value.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the label position for an axis.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_label_position.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartAxisLabelPosition, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 5)?;
    /// #     worksheet.write(1, 0, -30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(3, 0, -30)?;
    /// #     worksheet.write(4, 0, 5)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Set the axis label position to the bottom of the chart.
    ///     chart
    ///         .x_axis()
    ///         .set_label_position(ChartAxisLabelPosition::Low);
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_label_position.png">
    ///
    pub fn set_label_position(&mut self, position: ChartAxisLabelPosition) -> &mut ChartAxis {
        self.label_position = position;
        self
    }

    /// Set the axis position on or between the tick marks.
    ///
    /// In Excel there are two "Axis position" options for Category axes: "On
    /// tick marks" and "Between tick marks". This property has different
    /// default value for different chart types and isn't available for some
    /// chart types like Scatter. The `set_position_between_ticks()` method can
    /// be used to change the default value.
    ///
    /// Note, this property is only applicable to Category axes, see [Chart
    /// Value and Category Axes] for an explanation of the difference between
    /// Value and Category axes in Excel.
    ///
    /// [Chart Value and Category Axes]:
    ///     crate::chart#chart-value-and-category-axes
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. Its default value depends on the
    ///   chart type.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the axes data position relative to
    /// the tick marks. Notice that by setting the data columns "on" the tick
    /// the first and last columns are cut off by the plot area.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_position_between_ticks.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 5)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(3, 0, 30)?;
    /// #     worksheet.write(4, 0, 5)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Set the axes data position relative to the tick marks.
    ///     chart.x_axis().set_position_between_ticks(false);
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_axis_set_position_between_ticks.png">
    ///
    pub fn set_position_between_ticks(&mut self, enable: bool) -> &mut ChartAxis {
        self.position_between_ticks = enable;
        self
    }

    /// Set the interval of the axis labels.
    ///
    /// Set the interval of the axis labels for Category axes. This value is 1
    /// by default, i.e., there is one label shown per category. If needed it
    /// can be set to another value.
    ///
    /// Note, this property is only applicable to Category axes, see [Chart
    /// Value and Category Axes] for an explanation of the difference between
    /// Value and Category axes in Excel.
    ///
    /// # Parameters
    ///
    /// * `interval` - The interval for the category labels. The default is 1.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the label interval for an axis.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_label_interval.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 5)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(3, 0, 30)?;
    /// #     worksheet.write(4, 0, 5)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Set the interval for the axis labels.
    ///     chart.x_axis().set_label_interval(2);
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_label_interval.png">
    ///
    pub fn set_label_interval(&mut self, interval: u16) -> &mut ChartAxis {
        self.label_interval = interval;
        self
    }

    /// Set the interval of the axis ticks.
    ///
    /// Set the interval of the axis ticks for Category axes. This value is 1
    /// by default, i.e., there is one tick shown per category. If needed it
    /// can be set to another value.
    ///
    /// Note, this property is only applicable to Category axes, see [Chart
    /// Value and Category Axes] for an explanation of the difference between
    /// Value and Category axes in Excel.
    ///
    /// # Parameters
    ///
    /// * `interval` - The interval for the category ticks. The default is 1.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the tick interval for an axis.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_tick_interval.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 5)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(3, 0, 30)?;
    /// #     worksheet.write(4, 0, 5)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Set the interval for the axis ticks.
    ///     chart.x_axis().set_tick_interval(2);
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_tick_interval.png">
    ///
    pub fn set_tick_interval(&mut self, interval: u16) -> &mut ChartAxis {
        self.tick_interval = interval;
        self
    }

    /// Set the type of major tick for the axis.
    ///
    /// Excel supports 4 types of tick position:
    ///
    /// - None
    /// - Inside only
    /// - Outside only
    /// - Cross - inside and outside
    ///
    /// # Parameters
    ///
    /// * `tick_type` - a [`ChartAxisTickType`] enum value.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the tick types for chart axes.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_major_tick_type.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartAxisTickType, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 5)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(3, 0, 30)?;
    /// #     worksheet.write(4, 0, 5)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Set the tick types for chart axes.
    ///     chart
    ///         .x_axis()
    ///         .set_major_tick_type(ChartAxisTickType::None);
    ///     chart
    ///         .y_axis()
    ///         .set_major_tick_type(ChartAxisTickType::Outside)
    ///         .set_minor_tick_type(ChartAxisTickType::Cross);
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_major_tick_type.png">
    ///
    pub fn set_major_tick_type(&mut self, tick_type: ChartAxisTickType) -> &mut ChartAxis {
        self.major_tick_type = Some(tick_type);
        self
    }

    /// Set the type of minor tick for the axis.
    ///
    /// See [`set_major_tick_type()`](ChartAxis::set_major_tick_type) above for
    /// an explanation and example.
    ///
    /// # Parameters
    ///
    /// * `tick_type` - a [`ChartAxisTickType`] enum value.
    ///
    pub fn set_minor_tick_type(&mut self, tick_type: ChartAxisTickType) -> &mut ChartAxis {
        self.minor_tick_type = Some(tick_type);
        self
    }

    /// Set the log base of the axis range.
    ///
    /// This property is only applicable to value axes, see [Chart Value and
    /// Category Axes] for an explanation of the difference between Value and
    /// Category axes in Excel.
    ///
    /// [Chart Value and Category Axes]:
    ///     crate::chart#chart-value-and-category-axes
    ///
    /// # Parameters
    ///
    /// * `base` - The logarithm base. Should be >= 2.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating setting the logarithm base for chart axes.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_log_base.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 5)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(3, 0, 30)?;
    /// #     worksheet.write(4, 0, 5)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Change the default logarithm base for the Y-axis.
    ///     chart.y_axis().set_log_base(10);
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_axis_set_log_base.png">
    ///
    pub fn set_log_base(&mut self, base: u16) -> &mut ChartAxis {
        if base >= 2 {
            self.log_base = base;
        }
        self
    }

    /// Hide the chart axis.
    ///
    /// Hide the number or label section of the chart axis.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating hiding the chart axes.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_axis_set_hidden.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 5)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(3, 0, 30)?;
    /// #     worksheet.write(4, 0, 5)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Hide both axes.
    ///     chart.x_axis().set_hidden();
    ///     chart.y_axis().set_hidden();
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_hidden.png">
    ///
    pub fn set_hidden(&mut self) -> &mut ChartAxis {
        self.is_hidden = true;
        self
    }
}

#[derive(Clone, PartialEq)]
pub(crate) enum ChartAxisType {
    Category,
    Value,
    Date,
}

#[derive(Clone, Copy)]
pub(crate) enum ChartAxisPosition {
    Top,
    Bottom,
    Left,
    Right,
}

impl ChartAxisPosition {
    pub(crate) fn reverse(self) -> ChartAxisPosition {
        match self {
            Self::Top => Self::Bottom,
            Self::Left => Self::Right,
            Self::Right => Self::Left,
            Self::Bottom => Self::Top,
        }
    }
}

impl fmt::Display for ChartAxisPosition {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Top => write!(f, "t"),
            Self::Left => write!(f, "l"),
            Self::Right => write!(f, "r"),
            Self::Bottom => write!(f, "b"),
        }
    }
}

/// The `ChartAxisLabelPosition` enum defines the [`Chart`] axis label
/// positions.
///
/// This property is used in conjunction with
/// [`ChartAxis::set_label_position()`].
///
/// # Examples
///
/// A chart example demonstrating setting the label position for an axis.
///
/// ```
/// # // This code is available in examples/doc_chart_axis_set_label_position.rs
/// #
/// # use rust_xlsxwriter::{Chart, ChartAxisLabelPosition, ChartType, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the chart.
/// #     worksheet.write(0, 0, 5)?;
/// #     worksheet.write(1, 0, -30)?;
/// #     worksheet.write(2, 0, 40)?;
/// #     worksheet.write(3, 0, -30)?;
/// #     worksheet.write(4, 0, 5)?;
/// #
/// #     // Create a new chart.
///     let mut chart = Chart::new(ChartType::Column);
///
///     // Add a data series using Excel formula syntax to describe the range.
///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
///
///     // Set the axis label position to the bottom of the chart.
///     chart
///         .x_axis()
///         .set_label_position(ChartAxisLabelPosition::Low);
///
///     // Hide legend for clarity.
///     chart.legend().set_hidden();
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
/// src="https://rustxlsxwriter.github.io/images/chart_axis_set_label_position.png">
///
#[derive(Clone, Copy)]
pub enum ChartAxisLabelPosition {
    /// Position the axis labels next to the axis. The default.
    NextTo,

    /// Position the axis labels at the top of the chart, for horizontal axes,
    /// or to the right for vertical axes.
    High,

    /// Position the axis labels at the bottom of the chart, for horizontal
    /// axes, or to the left for vertical axes.
    Low,

    /// Turn off the the axis labels.
    None,
}

impl fmt::Display for ChartAxisLabelPosition {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Low => write!(f, "low"),
            Self::High => write!(f, "high"),
            Self::None => write!(f, "none"),
            Self::NextTo => write!(f, "nextTo"),
        }
    }
}

/// The `ChartAxisTickType` enum defines the [`Chart`] axis tick types.
///
/// Excel supports 4 types of tick position:
///
/// - None
/// - Inside only
/// - Outside only
/// - Cross - inside and outside
///
/// Used in conjunction with
/// [`set_major_tick_type()`](ChartAxis::set_major_tick_type) and
/// [`set_minor_tick_type()`](ChartAxis::set_minor_tick_type).
///
/// # Examples
///
/// A chart example demonstrating setting the tick types for chart axes.
///
/// ```
/// # // This code is available in examples/doc_chart_axis_set_major_tick_type.rs
/// #
/// # use rust_xlsxwriter::{Chart, ChartAxisTickType, ChartType, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the chart.
/// #     worksheet.write(0, 0, 5)?;
/// #     worksheet.write(1, 0, 30)?;
/// #     worksheet.write(2, 0, 40)?;
/// #     worksheet.write(3, 0, 30)?;
/// #     worksheet.write(4, 0, 5)?;
/// #
/// #     // Create a new chart.
///     let mut chart = Chart::new(ChartType::Column);
///
///     // Add a data series using Excel formula syntax to describe the range.
///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
///
///     // Set the tick types for chart axes.
///     chart
///         .x_axis()
///         .set_major_tick_type(ChartAxisTickType::None);
///     chart
///         .y_axis()
///         .set_major_tick_type(ChartAxisTickType::Outside)
///         .set_minor_tick_type(ChartAxisTickType::Cross);
///
///     // Hide legend for clarity.
///     chart.legend().set_hidden();
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
/// src="https://rustxlsxwriter.github.io/images/chart_axis_set_major_tick_type.png">
///
#[derive(Clone, Copy)]
pub enum ChartAxisTickType {
    /// No tick mark for the axis.
    None,

    /// The tick mark is inside the axis only.
    Inside,

    /// The tick mark is outside the axis only.
    Outside,

    /// The tick mark crosses inside and outside the axis.
    Cross,
}

impl fmt::Display for ChartAxisTickType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Cross => write!(f, "cross"),
            Self::Inside => write!(f, "in"),
            Self::Outside => write!(f, "out"),
        }
    }
}

/// The `ChartAxisDateUnitType` enum defines the [`Chart`] axis date unit types.
///
/// Define the unit type for the major or minor unit in a Chart Date axis.
///
#[derive(Clone, Copy)]
pub enum ChartAxisDateUnitType {
    /// The major or minor unit is expressed in days.
    Days,

    /// The major or minor unit is expressed in months.
    Months,

    /// The major or minor unit is expressed in years.
    Years,
}

impl fmt::Display for ChartAxisDateUnitType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Days => write!(f, "days"),
            Self::Years => write!(f, "years"),
            Self::Months => write!(f, "months"),
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

impl fmt::Display for ChartGrouping {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Stacked => write!(f, "stacked"),
            Self::Standard => write!(f, "standard"),
            Self::Clustered => write!(f, "clustered"),
            Self::PercentStacked => write!(f, "percentStacked"),
        }
    }
}

/// The `ChartAxisDisplayUnitType` enum defines the [`Chart`] axis date display
/// unit types.
///
/// Define the display unit type for chart axes such as "Thousands" or
/// "Millions".
///
#[derive(Clone, Copy, PartialEq)]
pub enum ChartAxisDisplayUnitType {
    /// Don't display any units for the axis values, the default.
    None,

    /// Display the axis values in units of Hundreds.
    Hundreds,

    /// Display the axis values in units of Thousands.
    Thousands,

    /// Display the axis values in units of Ten Thousands.
    TenThousands,

    /// Display the axis values in units of Hundred Thousands.
    HundredThousands,

    /// Display the axis values in units of Millions.
    Millions,

    /// Display the axis values in units of Ten Millions.
    TenMillions,

    /// Display the axis values in units of Hundred Millions.
    HundredMillions,

    /// Display the axis values in units of Billions.
    Billions,

    /// Display the axis values in units of Trillions.
    Trillions,
}

impl fmt::Display for ChartAxisDisplayUnitType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Hundreds => write!(f, "hundreds"),
            Self::Thousands => write!(f, "thousands"),
            Self::TenThousands => write!(f, "tenThousands"),
            Self::HundredThousands => write!(f, "hundredThousands"),
            Self::Millions => write!(f, "millions"),
            Self::TenMillions => write!(f, "tenMillions"),
            Self::HundredMillions => write!(f, "hundredMillions"),
            Self::Billions => write!(f, "billions"),
            Self::Trillions => write!(f, "trillions"),
        }
    }
}

// -----------------------------------------------------------------------
// ChartLegend
// -----------------------------------------------------------------------

/// The `ChartLegend` struct represents a chart legend.
///
/// The `ChartLegend` struct is a representation of a legend on an Excel chart.
/// The legend is a rectangular box that identifies the name and color of each
/// of the series in the chart.
///
/// `ChartLegend` can be used to configure properties of the chart legend and is
/// usually obtained via the [`chart.legend()`][Chart::legend] method.
///
/// It is used in conjunction with the [`Chart`] struct.
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
    pub(crate) font: Option<ChartFont>,
    deleted_entries: Vec<usize>,
}

impl ChartLegend {
    pub(crate) fn new() -> ChartLegend {
        ChartLegend {
            position: ChartLegendPosition::Right,
            hidden: false,
            has_overlay: false,
            format: ChartFormat::default(),
            font: None,
            deleted_entries: vec![],
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
    ///     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Hide the chart legend.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// The equivalent positions in `rust_xlsxwriter` charts are defined by
    /// [`ChartLegendPosition`]. The default chart position in Excel is to have
    /// the legend at the right.
    ///
    /// # Parameters
    ///
    /// `position` - the [`ChartLegendPosition`] position value.
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
    /// - [`ChartFormat::set_solid_fill()`]: Set the [`ChartSolidFill`] properties.
    /// - [`ChartFormat::set_pattern_fill()`]: Set the [`ChartPatternFill`] properties.
    /// - [`ChartFormat::set_gradient_fill()`]: Set the [`ChartGradientFill`] properties.
    /// - [`ChartFormat::set_no_fill()`]: Turn off the fill for the chart object.
    /// - [`ChartFormat::set_line()`]: Set the [`ChartLine`] properties.
    /// - [`ChartFormat::set_border()`]: Set the [`ChartBorder`] properties.
    ///   A synonym for [`ChartLine`] depending on context.
    /// - [`ChartFormat::set_no_line()`]: Turn off the line for the chart object.
    /// - [`ChartFormat::set_no_border()`]: Turn off the border for the chart object.
    /// - `set_no_border`: Turn off the border for the chart object.
    /// - `set_no_border`: Turn off the border for the chart object.
    ///
    /// # Parameters
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
    /// Set the font properties of a chart legend.
    ///
    /// Set the font properties of a chart legend using a [`ChartFont`]
    /// reference. Example font properties that can be set are:
    ///
    /// - [`ChartFont::set_bold()`]
    /// - [`ChartFont::set_italic()`]
    /// - [`ChartFont::set_name()`]
    /// - [`ChartFont::set_size()`]
    /// - [`ChartFont::set_rotation()`]
    ///
    /// See [`ChartFont`] for full details.
    ///
    /// # Parameters
    ///
    /// `font` - A [`ChartFont`] struct reference to represent the font
    /// properties.
    ///
    ///
    /// # Examples
    ///
    /// An example of setting the font for a chart legend.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_legend_set_font.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFont, ChartType, Workbook, XlsxError};
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
    ///     // Add a data series.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font.
    ///     chart
    ///         .legend()
    ///         .set_font(ChartFont::new().set_bold().set_color("#FF0000"));
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_legend_set_font.png">
    ///
    pub fn set_font(&mut self, font: &ChartFont) -> &mut ChartLegend {
        self.font = Some(font.clone());
        self
    }

    /// Delete/hide series names from the chart legend.
    ///
    /// The `delete_entries()` method deletes/hides one or more series names
    /// from the chart legend. This is sometimes required if there are a lot of
    /// secondary series and their names are cluttering the chart legend.
    ///
    /// The same effect can be accomplished using the
    /// [`ChartSeries::delete_from_legend()`] and
    /// [`ChartTrendline::delete_from_legend()`] methods. However, this method
    /// can be used for some edge cases such as Pie/Doughnut charts which
    /// display legend entries for each point in the series.
    ///
    /// Note, to hide all the names in the chart legend you should use the
    /// [`ChartLegend::set_hidden()`] method instead.
    ///
    /// # Parameters
    ///
    /// * `entries` - A slice ref of [`usize`] zero-indexed indices of the
    ///   series names to be hidden.
    ///
    /// # Examples
    ///
    /// A chart example demonstrating deleting/hiding a series name from the
    /// chart legend.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_legend_delete_entries.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 30)?;
    /// #     worksheet.write(1, 0, 20)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #     worksheet.write(0, 1, 10)?;
    /// #     worksheet.write(1, 1, 10)?;
    /// #     worksheet.write(2, 1, 10)?;
    /// #     worksheet.write(0, 2, 20)?;
    /// #     worksheet.write(1, 2, 15)?;
    /// #     worksheet.write(2, 2, 30)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add some data series.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///     chart.add_series().set_values("Sheet1!$B$1:$B$3");
    ///     chart.add_series().set_values("Sheet1!$C$1:$C$3");
    ///
    ///     // Delete the name of the second series (counted from zero) from the chart
    ///     // legend.
    ///     chart.legend().delete_entries(&[1]);
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 3, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_series_delete_from_legend.png">
    ///
    ///
    /// The default display without deleting the names from the legend would
    /// look like this:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_series_delete_from_legend2.png">
    ///
    pub fn delete_entries(&mut self, entries: &[usize]) -> &mut ChartLegend {
        self.deleted_entries = entries.to_vec();
        self
    }
}

/// The `ChartLegendPosition` enum defines the [`Chart`] legend positions.
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

impl fmt::Display for ChartLegendPosition {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Top => write!(f, "t"),
            Self::Left => write!(f, "l"),
            Self::Right => write!(f, "r"),
            Self::Bottom => write!(f, "b"),
            Self::TopRight => write!(f, "tr"),
        }
    }
}

/// The `ChartEmptyCells` enum defines the [`Chart`] empty cell options.
///
/// This enum defines the Excel chart options for handling empty cell in the
/// chart data ranges.
///
/// These options can be set using the
/// [`chart.show_empty_cells_as()`](Chart::show_empty_cells_as) method.
///
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ChartEmptyCells {
    /// Show empty cells in the chart as gaps. The default.
    Gaps,

    /// Show empty cells in the chart as zeroes.
    Zero,

    /// Show empty cells in the chart connected by a line to the previous point.
    Connected,
}

impl fmt::Display for ChartEmptyCells {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Gaps => write!(f, "gap"),
            Self::Zero => write!(f, "zero"),
            Self::Connected => write!(f, "span"),
        }
    }
}

// -----------------------------------------------------------------------
// ChartFormat
// -----------------------------------------------------------------------

#[derive(Clone, PartialEq)]
/// The `ChartFormat` struct represents formatting for various chart objects.
///
/// Excel uses a standard formatting dialog for the elements of a chart such as
/// data series, the plot area, the chart area, the legend or individual points.
/// It looks like this:
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_format_dialog.png">
///
/// The [`ChartFormat`] struct represents many of these format options and just
/// like Excel it offers a similar formatting interface for a number of the
/// chart sub-elements supported by `rust_xlsxwriter`.
///
/// It is used in conjunction with the [`Chart`] struct.
///
/// The [`ChartFormat`] struct is generally passed to the `set_format()` method
/// of a chart element. It supports several chart formatting elements such as:
///
/// - [`ChartFormat::set_solid_fill()`]: Set the [`ChartSolidFill`] properties.
/// - [`ChartFormat::set_pattern_fill()`]: Set the [`ChartPatternFill`]
///   properties.
/// - [`ChartFormat::set_gradient_fill()`]: Set the [`ChartGradientFill`]
///   properties.
/// - [`ChartFormat::set_no_fill()`]: Turn off the fill for the chart object.
/// - [`ChartFormat::set_line()`]: Set the [`ChartLine`] properties.
/// - [`ChartFormat::set_border()`]: Set the [`ChartBorder`] properties. A
///   synonym for [`ChartLine`] depending on context.
/// - [`ChartFormat::set_no_line()`]: Turn off the line for the chart object.
/// - [`ChartFormat::set_no_border()`]: Turn off the border for the chart
///   object.
///
/// # Examples
///
/// An example of accessing the [`ChartFormat`] for data series in a chart and
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
    gradient_fill: Option<ChartGradientFill>,
}

impl Default for ChartFormat {
    fn default() -> Self {
        Self::new()
    }
}

impl ChartFormat {
    /// Create a new `ChartFormat` instance to set formatting for a chart element.
    ///
    pub fn new() -> ChartFormat {
        ChartFormat {
            no_fill: false,
            no_line: false,
            line: None,
            solid_fill: None,
            pattern_fill: None,
            gradient_fill: None,
        }
    }

    /// Set the line formatting for a chart element.
    ///
    /// See the [`ChartLine`] struct for details on the line properties that can
    /// be set.
    ///
    /// # Parameters
    ///
    /// * `line` - A [`ChartLine`] struct reference.
    ///
    pub fn set_line(&mut self, line: &ChartLine) -> &mut ChartFormat {
        self.line = Some(line.clone());
        self
    }

    /// Set the border formatting for a chart element.
    ///
    /// See the [`ChartLine`] struct for details on the border properties that
    /// can be set. As a syntactic shortcut you can use the type alias
    /// [`ChartBorder`] instead
    /// of `ChartLine`.
    ///
    /// # Parameters
    ///
    /// * `line` - A [`ChartLine`] struct reference.
    ///
    /// # Examples
    ///
    /// An example of formatting the border in a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_border_formatting.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartBorder, ChartFormat, ChartLineDashType, ChartType, Workbook, XlsxError,
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
    ///             ChartFormat::new()
    ///                 .set_border(
    ///                     ChartBorder::new()
    ///                         .set_color("#FF9900")
    ///                         .set_width(5.25)
    ///                         .set_dash_type(ChartLineDashType::SquareDot)
    ///                         .set_transparency(70),
    ///                 )
    ///                 .set_no_fill(),
    ///         );
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_border_formatting.png">
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
    /// # use rust_xlsxwriter::{Chart, ChartFormat, ChartLine, ChartType, Workbook, Color, XlsxError};
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
    ///                 .set_border(ChartLine::new().set_color(Color::Black))
    ///                 .set_no_fill(),
    ///         );
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// # Parameters
    ///
    /// * `fill` - A [`ChartSolidFill`] struct reference.
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
    /// # Parameters
    ///
    /// * `fill` - A [`ChartPatternFill`] struct reference.
    ///
    /// # Examples
    ///
    /// An example of setting a pattern fill for a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_pattern_fill.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartFormat, ChartPatternFill, ChartPatternFillType, ChartType, Workbook, Color,
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
    ///                     .set_background_color(Color::Yellow)
    ///                     .set_foreground_color(Color::Red),
    ///             ),
    ///         );
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill.png">
    ///
    pub fn set_pattern_fill(&mut self, fill: &ChartPatternFill) -> &mut ChartFormat {
        self.pattern_fill = Some(fill.clone());
        self
    }

    /// Set the gradient fill formatting for a chart element.
    ///
    /// See the [`ChartGradientFill`] struct for details on the gradient fill
    /// properties that can be set.
    ///
    /// # Parameters
    ///
    /// * `fill` - A [`ChartGradientFill`] struct reference.
    ///
    /// # Examples
    ///
    /// An example of setting a gradient fill for a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_format_set_gradient_fill.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartFormat, ChartGradientFill, ChartGradientStop, ChartType, Workbook, XlsxError,
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
    ///         .set_format(ChartFormat::new().set_gradient_fill(
    ///             ChartGradientFill::new().set_gradient_stops(&[
    ///                 ChartGradientStop::new("#963735", 0),
    ///                 ChartGradientStop::new("#F1DCDB", 100),
    ///             ]),
    ///         ));
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_gradient_fill.png">
    ///
    pub fn set_gradient_fill(&mut self, fill: &ChartGradientFill) -> &mut ChartFormat {
        self.gradient_fill = Some(fill.clone());
        self
    }

    // Check if formatting has been set for the struct.
    fn has_formatting(&self) -> bool {
        self.line.is_some()
            || self.solid_fill.is_some()
            || self.pattern_fill.is_some()
            || self.gradient_fill.is_some()
            || self.no_fill
            || self.no_line
    }
}

/// The `ChartLine` struct represents a chart line/border.
///
/// The [`ChartLine`] struct represents the formatting properties for a line or
/// border for a Chart element. It is a sub property of the [`ChartFormat`]
/// struct and is used with the [`ChartFormat::set_line()`] or
/// [`ChartFormat::set_border()`] methods.
///
/// Excel uses the element names "Line" and "Border" depending on the context.
/// For a Line chart the line is represented by a line property but for a Column
/// chart the line becomes the border. Both of these share the same properties
/// and are both represented in `rust_xlsxwriter` by the [`ChartLine`] struct.
///
/// As a syntactic shortcut you can use the type alias [`ChartBorder`] instead
/// of `ChartLine`.
///
/// It is used in conjunction with the [`Chart`] struct.
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
/// <img
/// src="https://rustxlsxwriter.github.io/images/chart_line_formatting.png">
///
#[derive(Clone, PartialEq)]
pub struct ChartLine {
    color: Color,
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
            color: Color::Default,
            width: None,
            transparency: 0,
            dash_type: ChartLineDashType::Solid,
            hidden: false,
        }
    }

    /// Set the color of a line/border.
    ///
    /// # Parameters
    ///
    /// * `color` - The color property defined by a [`Color`] enum value or
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
    /// # Parameters
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
    /// # Parameters
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
    ///             ChartFormat::new()
    ///                 .set_line(ChartLine::new()
    ///                 .set_dash_type(ChartLineDashType::DashDot)),
    ///         );
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// # Parameters
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_set_transparency.png">
    ///
    pub fn set_transparency(&mut self, transparency: u8) -> &mut ChartLine {
        if transparency <= 100 {
            self.transparency = transparency;
        }

        self
    }

    /// Set the chart line as hidden.
    ///
    /// The method is sometimes required to turn off a default line type in
    /// order to highlight some other element such as the line markers.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off (not hidden) by
    ///   default.
    ///
    pub fn set_hidden(&mut self, enable: bool) -> &mut ChartLine {
        self.hidden = enable;
        self
    }
}

/// A type to represent a Chart border. It can be used interchangeably with
/// [`ChartLine`].
///
/// Excel uses the chart element names "Line" and "Border" depending on the
/// context. For a Line chart the line is represented by a line property but for
/// a Column chart the line becomes the border. Both of these share the same
/// properties and are both represented in `rust_xlsxwriter` by the
/// [`ChartLine`] struct.
///
/// The `ChartBorder` type is a type alias of [`ChartLine`] for use as a
/// syntactic shortcut where you would expect to write `ChartBorder` instead of
/// `ChartLine`.
///
/// # Examples
///
/// An example of formatting the border in a chart element.
///
/// ```
/// # // This code is available in examples/doc_chart_border_formatting.rs
/// #
/// # use rust_xlsxwriter::{
/// #     Chart, ChartBorder, ChartFormat, ChartLineDashType, ChartType, Workbook, XlsxError,
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
///             ChartFormat::new()
///                 .set_border(
///                     ChartBorder::new()
///                         .set_color("#FF9900")
///                         .set_width(5.25)
///                         .set_dash_type(ChartLineDashType::SquareDot)
///                         .set_transparency(70),
///                 )
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
/// <img
/// src="https://rustxlsxwriter.github.io/images/chart_border_formatting.png">
///
pub type ChartBorder = ChartLine;

/// The `ChartSolidFill` struct represents a the solid fill for a chart element.
///
/// The [`ChartSolidFill`] struct represents the formatting properties for the
/// solid fill of a Chart element. In Excel a solid fill is a single color fill
/// without a pattern or gradient.
///
/// `ChartSolidFill` is a sub property of the [`ChartFormat`] struct and is used
/// with the [`ChartFormat::set_solid_fill()`] method.
///
/// It is used in conjunction with the [`Chart`] struct.
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
#[derive(Clone, PartialEq)]
pub struct ChartSolidFill {
    color: Color,
    transparency: u8,
}

impl ChartSolidFill {
    /// Create a new `ChartSolidFill` object to represent a Chart solid fill.
    ///
    #[allow(clippy::new_without_default)]
    pub fn new() -> ChartSolidFill {
        ChartSolidFill {
            color: Color::Default,
            transparency: 0,
        }
    }

    /// Set the color of a solid fill.
    ///
    /// # Parameters
    ///
    /// * `color` - The color property defined by a [`Color`] enum value or
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
    /// # Parameters
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_solid_fill.png">
    ///
    pub fn set_transparency(&mut self, transparency: u8) -> &mut ChartSolidFill {
        if transparency <= 100 {
            self.transparency = transparency;
        }

        self
    }
}

/// The `ChartPatternFill` struct represents a the pattern fill for a chart
/// element.
///
/// The [`ChartPatternFill`] struct represents the formatting properties for the
/// pattern fill of a Chart element. In Excel a pattern fill is comprised of a
/// simple pixelated pattern and background and foreground colors
///
/// `ChartPatternFill` is a sub property of the [`ChartFormat`] struct and is
/// used with the [`ChartFormat::set_pattern_fill()`] method.
///
/// It is used in conjunction with the [`Chart`] struct.
///
/// # Examples
///
/// An example of setting a pattern fill for a chart element.
///
/// ```
/// # // This code is available in examples/doc_chart_pattern_fill.rs
/// #
/// # use rust_xlsxwriter::{
/// #     Chart, ChartFormat, ChartPatternFill, ChartPatternFillType, ChartType, Workbook, Color, XlsxError,
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
///                     .set_background_color(Color::Yellow)
///                     .set_foreground_color(Color::Red),
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
#[derive(Clone, PartialEq)]
pub struct ChartPatternFill {
    background_color: Color,
    foreground_color: Color,
    pattern: ChartPatternFillType,
}

impl ChartPatternFill {
    /// Create a new `ChartPatternFill` object to represent a Chart pattern fill.
    ///
    #[allow(clippy::new_without_default)]
    pub fn new() -> ChartPatternFill {
        ChartPatternFill {
            background_color: Color::Default,
            foreground_color: Color::Default,
            pattern: ChartPatternFillType::Dotted5Percent,
        }
    }

    /// Set the pattern of a Chart pattern fill element.
    ///
    /// See the example above.
    ///
    /// # Parameters
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
    /// # Parameters
    ///
    /// * `color` - The color property defined by a [`Color`] enum value or
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
    /// #     Chart, ChartFormat, ChartPatternFill, ChartPatternFillType, ChartType, Workbook, Color,
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
    ///                     .set_background_color(Color::Yellow)
    ///                     .set_foreground_color(Color::Red),
    ///             ),
    ///         );
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// # Parameters
    ///
    /// * `color` - The color property defined by a [`Color`] enum value or
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

/// The `ChartLineDashType` enum defines the [`Chart`] line dash types.
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

impl fmt::Display for ChartLineDashType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Dash => write!(f, "dash"),
            Self::Solid => write!(f, "solid"),
            Self::DashDot => write!(f, "dashDot"),
            Self::LongDash => write!(f, "lgDash"),
            Self::RoundDot => write!(f, "sysDot"),
            Self::SquareDot => write!(f, "sysDash"),
            Self::LongDashDot => write!(f, "lgDashDot"),
            Self::LongDashDotDot => write!(f, "lgDashDotDot"),
        }
    }
}

/// The `ChartPatternFillType` enum defines the [`Chart`] pattern fill types.
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

impl fmt::Display for ChartPatternFillType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Wave => write!(f, "wave"),
            Self::Weave => write!(f, "weave"),
            Self::Plaid => write!(f, "plaid"),
            Self::Divot => write!(f, "divot"),
            Self::Zigzag => write!(f, "zigZag"),
            Self::Sphere => write!(f, "sphere"),
            Self::Shingle => write!(f, "shingle"),
            Self::Trellis => write!(f, "trellis"),
            Self::SmallGrid => write!(f, "smGrid"),
            Self::LargeGrid => write!(f, "lgGrid"),
            Self::DottedGrid => write!(f, "dotGrid"),
            Self::DottedDiamond => write!(f, "dotDmnd"),
            Self::DiagonalBrick => write!(f, "diagBrick"),
            Self::LargeConfetti => write!(f, "lgConfetti"),
            Self::SmallConfetti => write!(f, "smConfetti"),
            Self::Dotted5Percent => write!(f, "pct5"),
            Self::Dotted10Percent => write!(f, "pct10"),
            Self::Dotted20Percent => write!(f, "pct20"),
            Self::Dotted25Percent => write!(f, "pct25"),
            Self::Dotted30Percent => write!(f, "pct30"),
            Self::Dotted40Percent => write!(f, "pct40"),
            Self::Dotted50Percent => write!(f, "pct50"),
            Self::Dotted60Percent => write!(f, "pct60"),
            Self::Dotted70Percent => write!(f, "pct70"),
            Self::Dotted75Percent => write!(f, "pct75"),
            Self::Dotted80Percent => write!(f, "pct80"),
            Self::Dotted90Percent => write!(f, "pct90"),
            Self::HorizontalBrick => write!(f, "horzBrick"),
            Self::SolidDiamondGrid => write!(f, "solidDmnd"),
            Self::SmallCheckerboard => write!(f, "smCheck"),
            Self::LargeCheckerboard => write!(f, "lgCheck"),
            Self::StripesBackslashes => write!(f, "dashDnDiag"),
            Self::VerticalStripesDark => write!(f, "dkVert"),
            Self::OutlinedDiamondGrid => write!(f, "openDmnd"),
            Self::VerticalStripesLight => write!(f, "ltVert"),
            Self::HorizontalStripesDark => write!(f, "dkHorz"),
            Self::StripesForwardSlashes => write!(f, "dashUpDiag"),
            Self::VerticalStripesNarrow => write!(f, "narVert"),
            Self::HorizontalStripesLight => write!(f, "ltHorz"),
            Self::HorizontalStripesNarrow => write!(f, "narHorz"),
            Self::DiagonalStripesDarkUpwards => write!(f, "dkUpDiag"),
            Self::DiagonalStripesWideUpwards => write!(f, "wdUpDiag"),
            Self::VerticalStripesAlternating => write!(f, "dashVert"),
            Self::DiagonalStripesLightUpwards => write!(f, "ltUpDiag"),
            Self::DiagonalStripesDarkDownwards => write!(f, "dkDnDiag"),
            Self::DiagonalStripesWideDownwards => write!(f, "wdDnDiag"),
            Self::HorizontalStripesAlternating => write!(f, "dashHorz"),
            Self::DiagonalStripesLightDownwards => write!(f, "ltDnDiag"),
        }
    }
}

// -----------------------------------------------------------------------
// ChartFont
// -----------------------------------------------------------------------

#[derive(Clone, PartialEq)]
/// The `ChartFont` struct represents the font format for various chart objects.
///
/// Excel uses a standard font dialog for text elements of a chart such as the
/// chart title or axes data labels. It looks like this:
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_font_dialog.png">
///
/// The [`ChartFont`] struct represents many of these font options such as font
/// type, size, color and properties such as bold and italic. It is generally
/// used in conjunction with a `set_font()` method for a chart element.
///
/// It is used in conjunction with the [`Chart`] struct.
///
/// # Examples
///
/// An example of setting the font for a chart element.
///
/// ```
/// # // This code is available in examples/doc_chart_font.rs
/// #
/// # use rust_xlsxwriter::{Chart, ChartFont, ChartType, Workbook, XlsxError};
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
///     // Add a data series.
///     chart.add_series().set_values("Sheet1!$A$1:$A$6");
///
///     // Set the font for an axis.
///     chart.x_axis().set_font(
///         ChartFont::new()
///             .set_bold()
///             .set_italic()
///             .set_name("Consolas")
///             .set_color("#FF0000"),
///     );
///
///     // Hide legend for clarity.
///     chart.legend().set_hidden();
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
/// <img src="https://rustxlsxwriter.github.io/images/chart_font.png">
///
pub struct ChartFont {
    // Chart/axis titles have a default bold font so we need to handle that as
    // an option.
    pub(crate) bold: Option<bool>,
    pub(crate) has_default_bold: bool,

    pub(crate) italic: bool,
    pub(crate) underline: bool,
    pub(crate) name: String,
    pub(crate) size: f64,
    pub(crate) color: Color,
    pub(crate) strikethrough: bool,
    pub(crate) pitch_family: u8,
    pub(crate) character_set: u8,
    pub(crate) rotation: Option<i16>,
    pub(crate) has_baseline: bool,
    pub(crate) right_to_left: Option<bool>,
}

impl Default for ChartFont {
    fn default() -> Self {
        Self::new()
    }
}

impl ChartFont {
    /// Create a new `ChartFont` object to represent a Chart font.
    ///
    pub fn new() -> ChartFont {
        ChartFont {
            bold: None,
            italic: false,
            underline: false,
            name: String::new(),
            size: 0.0,
            color: Color::Default,
            strikethrough: false,
            pitch_family: 0,
            character_set: 0,
            rotation: None,
            has_baseline: false,
            has_default_bold: false,
            right_to_left: None,
        }
    }

    /// Set the bold property for the font of a chart element.
    ///
    /// # Examples
    ///
    /// An example of setting the bold property for the font in a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_font_set_bold.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFont, ChartType, Workbook, XlsxError};
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
    ///     // Add a data series.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font for an axis.
    ///     chart.x_axis().set_font(ChartFont::new().set_bold());
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_font_set_bold.png">
    ///
    pub fn set_bold(&mut self) -> &mut ChartFont {
        self.bold = Some(true);
        self
    }

    /// Set the italic property for the font of a chart element.
    ///
    /// # Examples
    ///
    /// An example of setting the italic property for the font in a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_font_set_italic.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFont, ChartType, Workbook, XlsxError};
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
    ///     // Add a data series.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font for an axis.
    ///     chart.x_axis().set_font(ChartFont::new().set_italic());
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_font_set_italic.png">
    ///
    pub fn set_italic(&mut self) -> &mut ChartFont {
        self.italic = true;
        self
    }

    /// Set the color property for the font of a chart element.
    ///
    /// # Parameters
    ///
    /// * `color` - The font color property defined by a [`Color`] enum
    ///   value.
    ///
    /// # Examples
    ///
    /// An example of setting the color property for the font in a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_font_set_color.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFont, ChartType, Workbook, XlsxError};
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
    ///     // Add a data series.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font for an axis.
    ///     chart.x_axis().set_font(ChartFont::new().set_color("#FF0000"));
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_font_set_color.png">
    ///
    pub fn set_color<T>(&mut self, color: T) -> &mut ChartFont
    where
        T: IntoColor,
    {
        let color = color.new_color();
        if color.is_valid() {
            self.color = color;
        }

        self
    }

    /// Set the chart font name property.
    ///
    /// Set the name/type of a font for a chart element.
    ///
    /// # Parameters
    ///
    /// * `font_name` - The font name property.
    ///
    ///
    /// # Examples
    ///
    /// An example of setting the font name property for the font in a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_font_set_name.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFont, ChartType, Workbook, XlsxError};
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
    ///     // Add a data series.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font for an axis.
    ///     chart
    ///         .x_axis()
    ///         .set_font(ChartFont::new().set_name("American Typewriter"));
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_font_set_name.png">
    ///
    pub fn set_name(&mut self, font_name: impl Into<String>) -> &mut ChartFont {
        self.name = font_name.into();
        self
    }

    /// Set the size property for the font of a chart element.
    ///
    /// # Parameters
    ///
    /// * `font_size` - The font size property.
    ///
    /// # Examples
    ///
    /// An example of setting the font size property for the font in a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_font_set_size.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFont, ChartType, Workbook, XlsxError};
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
    ///     // Add a data series.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font for an axis.
    ///     chart.x_axis().set_font(ChartFont::new().set_size(20));
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_font_set_size.png">
    ///
    pub fn set_size<T>(&mut self, font_size: T) -> &mut ChartFont
    where
        T: Into<f64>,
    {
        self.size = font_size.into() * 100.0;
        self
    }

    /// Set the text rotation property for the font of a chart element.
    ///
    /// Set the rotation angle of the text in a cell. The rotation can be any
    /// angle in the range -90 to 90 degrees, or 270-271 to indicate text where
    /// the letters run from top to bottom, see below.
    ///
    /// # Parameters
    ///
    /// * `rotation` - The rotation angle in the range `-90 <= rotation <= 90`.
    ///   Two special case values are supported:
    ///   - `270`: Stacked text, where the text runs from top to bottom.
    ///   - `271`: A special variant of stacked text for East Asian fonts.
    ///
    /// # Examples
    ///
    /// An example of setting the font text rotation for the font in a chart
    /// element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_font_set_rotation.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartFont, ChartType, Workbook, XlsxError};
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
    ///     // Add a data series.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font for an axis.
    ///     chart.x_axis().set_font(ChartFont::new().set_rotation(45));
    ///
    ///     // Hide legend for clarity.
    ///     chart.legend().set_hidden();
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_font_set_rotation.png">
    ///
    pub fn set_rotation(&mut self, rotation: i16) -> &mut ChartFont {
        match rotation {
            270..=271 | -90..=90 => self.rotation = Some(rotation),
            _ => eprintln!("Rotation '{rotation}' outside range: -90 <= angle <= 90."),
        }

        self
    }

    /// Set the underline property for the font of a chart element.
    ///
    /// The default underline type is the only type supported.
    ///
    pub fn set_underline(&mut self) -> &mut ChartFont {
        self.underline = true;
        self
    }

    /// Set the strikethrough property for the font of a chart element.
    ///
    pub fn set_strikethrough(&mut self) -> &mut ChartFont {
        self.strikethrough = true;
        self
    }

    /// Unset the bold property for a font.
    ///
    /// Some chart elements such as titles have a default bold property in
    /// Excel. This method can be used to turn it off.
    ///
    pub fn unset_bold(&mut self) -> &mut ChartFont {
        self.bold = Some(false);
        self
    }

    /// Display the chart font from right to left for some language support.
    ///
    /// See
    /// [`Worksheet::set_right_to_left()`](crate::Worksheet::set_right_to_left)
    /// for details.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn set_right_to_left(&mut self, enable: bool) -> &mut ChartFont {
        self.right_to_left = Some(enable);
        self
    }

    /// Set the pitch family property for the font of a chart element.
    ///
    /// This function is implemented for completeness but is rarely used in
    /// practice.
    ///
    /// # Parameters
    ///
    /// * `family` - The font family property.
    ///
    pub fn set_pitch_family(&mut self, family: u8) -> &mut ChartFont {
        self.pitch_family = family;
        self
    }

    /// Set the character set property for the font of a chart element.
    ///
    /// Set the font character set. This function is implemented for
    /// completeness but is rarely required in practice.
    ///
    /// # Parameters
    ///
    /// * `character_set` - The font character set property.
    ///
    pub fn set_character_set(&mut self, character_set: u8) -> &mut ChartFont {
        self.character_set = character_set;
        self
    }

    // Internal check for font properties that need a sub-element.
    pub(crate) fn is_latin(&self) -> bool {
        !self.name.is_empty() || self.pitch_family > 0 || self.character_set > 0
    }
}

// -----------------------------------------------------------------------
// ChartTrendline
// -----------------------------------------------------------------------

/// The `ChartTrendline` struct represents a trendline for a chart series.
///
/// Excel allows you to add a trendline to a data series that represents the
/// trend or regression of the data using different types of fit. The
/// `ChartTrendline` struct represents the options of Excel trendlines and can
/// be added to a series via the [`ChartSeries::set_trendline()`] method.
///
/// <img src="https://rustxlsxwriter.github.io/images/trendline_options.png">
///
/// # Examples
///
/// An example of adding a trendline to a chart data series. The options used
/// are shown in the image above.
///
/// ```
/// # // This code is available in examples/doc_chart_trendline_intro.rs
/// #
/// # use rust_xlsxwriter::{Chart, ChartTrendline, ChartTrendlineType, ChartType, Workbook, XlsxError};
/// #
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///     let worksheet = workbook.add_worksheet();
///
///     // Add some data for the chart.
///     worksheet.write(0, 0, 11.1)?;
///     worksheet.write(1, 0, 18.8)?;
///     worksheet.write(2, 0, 33.2)?;
///     worksheet.write(3, 0, 37.5)?;
///     worksheet.write(4, 0, 52.1)?;
///     worksheet.write(5, 0, 58.9)?;
///
///     // Create a trendline.
///     let mut trendline = ChartTrendline::new();
///     trendline
///         .set_type(ChartTrendlineType::Linear)
///         .display_equation(true)
///         .display_r_squared(true);
///
///     // Create a new chart.
///     let mut chart = Chart::new(ChartType::Line);
///
///     // Add a data series with a trendline.
///     chart
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_trendline(&trendline);
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
/// <img
/// src="https://rustxlsxwriter.github.io/images/chart_trendline_intro.png">
///
#[derive(Clone)]
pub struct ChartTrendline {
    name: String,
    trend_type: ChartTrendlineType,
    format: ChartFormat,
    label_format: ChartFormat,
    label_font: Option<ChartFont>,
    forward_period: f64,
    backward_period: f64,
    display_equation: bool,
    display_r_squared: bool,
    intercept: Option<f64>,
    delete_from_legend: bool,
}

impl ChartTrendline {
    /// Create a new `ChartTrendline` object to represent a Chart series trendline.
    #[allow(clippy::new_without_default)]
    pub fn new() -> ChartTrendline {
        ChartTrendline {
            name: String::new(),
            trend_type: ChartTrendlineType::None,
            format: ChartFormat::default(),
            label_format: ChartFormat::default(),
            label_font: None,
            forward_period: 0.0,
            backward_period: 0.0,
            display_r_squared: false,
            display_equation: false,
            intercept: None,
            delete_from_legend: false,
        }
    }

    /// Set the type of the Chart series trendlines.
    ///
    /// Set the trendline type to one of the Excel allowable types represented
    /// by the [`ChartTrendlineType`] enum.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/trendline_types.png">
    ///
    /// # Parameters
    ///
    /// * `trend` - A [`ChartTrendlineType`] enum reference.
    ///
    /// # Examples
    ///
    /// An example of adding a trendline to a chart data series. Demonstrates
    /// setting the polynomial trendline type.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_trendline_set_type.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartTrendline, ChartTrendlineType, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 11.1)?;
    /// #     worksheet.write(1, 0, 18.8)?;
    /// #     worksheet.write(2, 0, 33.2)?;
    /// #     worksheet.write(3, 0, 37.5)?;
    /// #     worksheet.write(4, 0, 52.1)?;
    /// #     worksheet.write(5, 0, 58.9)?;
    /// #
    ///     // Create a trendline.
    ///     let mut trendline = ChartTrendline::new();
    ///     trendline
    ///         .set_type(ChartTrendlineType::Polynomial(3))
    ///         .display_equation(true);
    ///
    ///     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with a trendline.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_trendline(&trendline);
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_trendline_set_type.png">
    ///
    pub fn set_type(&mut self, trend: ChartTrendlineType) -> &mut ChartTrendline {
        self.trend_type = trend;
        self
    }

    /// Set the formatting properties for a chart trendline.
    ///
    /// Set the formatting properties for a chart trendline via a
    /// [`ChartFormat`] object or a sub struct that implements
    /// [`IntoChartFormat`].
    ///
    /// The formatting that can be applied via a [`ChartFormat`] object are:
    /// - [`ChartFormat::set_solid_fill()`]: Set the [`ChartSolidFill`] properties.
    /// - [`ChartFormat::set_pattern_fill()`]: Set the [`ChartPatternFill`] properties.
    /// - [`ChartFormat::set_gradient_fill()`]: Set the [`ChartGradientFill`] properties.
    /// - [`ChartFormat::set_no_fill()`]: Turn off the fill for the chart object.
    /// - [`ChartFormat::set_line()`]: Set the [`ChartLine`] properties.
    /// - [`ChartFormat::set_border()`]: Set the [`ChartBorder`] properties.
    ///   A synonym for [`ChartLine`] depending on context.
    /// - [`ChartFormat::set_no_line()`]: Turn off the line for the chart object.
    /// - [`ChartFormat::set_no_border()`]: Turn off the border for the chart object.
    /// - `set_no_border`: Turn off the border for the chart object.
    ///
    /// # Parameters
    ///
    /// `format`: A [`ChartFormat`] struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// [`IntoChartFormat`] for details.
    ///
    /// # Examples
    ///
    /// An example of adding a trendline to a chart data series with formatting.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_trendline_set_format.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartLine, ChartLineDashType, ChartTrendline, ChartTrendlineType, ChartType, Color,
    /// #     Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 11.1)?;
    /// #     worksheet.write(1, 0, 18.8)?;
    /// #     worksheet.write(2, 0, 33.2)?;
    /// #     worksheet.write(3, 0, 37.5)?;
    /// #     worksheet.write(4, 0, 52.1)?;
    /// #     worksheet.write(5, 0, 58.9)?;
    /// #
    ///     // Create a trendline.
    ///     let mut trendline = ChartTrendline::new();
    ///     trendline.set_type(ChartTrendlineType::Linear).set_format(
    ///         ChartLine::new()
    ///             .set_color(Color::Red)
    ///             .set_width(1)
    ///             .set_dash_type(ChartLineDashType::LongDash),
    ///     );
    ///
    ///     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with a trendline.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_trendline(&trendline);
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_trendline_set_format.png">
    ///
    pub fn set_format<T>(&mut self, format: T) -> &mut ChartTrendline
    where
        T: IntoChartFormat,
    {
        self.format = format.new_chart_format();
        self
    }

    /// Set the formatting properties for a chart trendline label.
    ///
    /// Set the formatting properties for a chart trendline label via a
    /// [`ChartFormat`] object or a sub struct that implements
    /// [`IntoChartFormat`]. The label is displayed when you use the
    /// [`display_equation()`](ChartTrendline::display_equation) or
    /// [`display_r_squared()`](ChartTrendline::display_equation) methods.
    ///
    /// The formatting that can be applied via a [`ChartFormat`] object are:
    ///
    /// - [`ChartFormat::set_solid_fill()`]: Set the [`ChartSolidFill`] properties.
    /// - [`ChartFormat::set_pattern_fill()`]: Set the [`ChartPatternFill`] properties.
    /// - [`ChartFormat::set_gradient_fill()`]: Set the [`ChartGradientFill`] properties.
    /// - [`ChartFormat::set_no_fill()`]: Turn off the fill for the chart object.
    /// - [`ChartFormat::set_line()`]: Set the [`ChartLine`] properties.
    /// - [`ChartFormat::set_border()`]: Set the [`ChartBorder`] properties.
    ///   A synonym for [`ChartLine`] depending on context.
    /// - [`ChartFormat::set_no_line()`]: Turn off the line for the chart object.
    /// - [`ChartFormat::set_no_border()`]: Turn off the border for the chart object.
    ///
    /// # Parameters
    ///
    /// `format`: A [`ChartFormat`] struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// [`IntoChartFormat`] for details.
    ///
    ///
    /// # Examples
    ///
    /// An example of adding a trendline to a chart data series and adding
    /// formatting to the trendline data label.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_trendline_set_label_format.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartFormat, ChartLine, ChartSolidFill, ChartTrendline, ChartTrendlineType, ChartType,
    /// #     Color, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 11.1)?;
    /// #     worksheet.write(1, 0, 18.8)?;
    /// #     worksheet.write(2, 0, 33.2)?;
    /// #     worksheet.write(3, 0, 37.5)?;
    /// #     worksheet.write(4, 0, 52.1)?;
    /// #     worksheet.write(5, 0, 58.9)?;
    /// #
    ///     // Create a trendline.
    ///     let mut trendline = ChartTrendline::new();
    ///     trendline
    ///         .set_type(ChartTrendlineType::Linear)
    ///         .display_equation(true)
    ///         .set_label_format(
    ///             ChartFormat::new()
    ///                 .set_solid_fill(ChartSolidFill::new().set_color(Color::Yellow))
    ///                 .set_border(ChartLine::new().set_color(Color::Red)),
    ///         );
    ///
    ///     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with a trendline.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_trendline(&trendline);
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_trendline_set_label_format.png">
    ///
    pub fn set_label_format<T>(&mut self, format: T) -> &mut ChartTrendline
    where
        T: IntoChartFormat,
    {
        self.label_format = format.new_chart_format();
        self
    }

    /// Set the font properties of a chart trendline label.
    ///
    /// Set the font properties of a chart trendline label using a [`ChartFont`]
    /// reference. The label is displayed when you use the
    /// [`display_equation()`](ChartTrendline::display_equation) or
    /// [`display_r_squared()`](ChartTrendline::display_equation) methods.
    ///
    /// Example font properties that can be set are:
    ///
    /// - [`ChartFont::set_bold()`]
    /// - [`ChartFont::set_italic()`]
    /// - [`ChartFont::set_name()`]
    /// - [`ChartFont::set_size()`]
    /// - [`ChartFont::set_rotation()`]
    ///
    /// See [`ChartFont`] for full details.
    ///
    /// # Parameters
    ///
    /// `font`: A [`ChartFont`] struct reference to represent the font
    /// properties.
    ///
    pub fn set_label_font(&mut self, font: &ChartFont) -> &mut ChartTrendline {
        let mut font = font.clone();
        font.has_baseline = true;
        self.label_font = Some(font);
        self
    }

    /// Set the name for a chart trendline.
    ///
    /// Set a custom name for a the trendline when it is displayed in the chart
    /// legend.
    ///
    /// # Parameters
    ///
    /// * `name` - The custom string to name the trendline in the chart legend.
    ///
    /// # Examples
    ///
    /// An example of adding a trendline to a chart data series with a custom
    /// name.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_trendline_set_name.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartTrendline, ChartTrendlineType, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 11.1)?;
    /// #     worksheet.write(1, 0, 18.8)?;
    /// #     worksheet.write(2, 0, 33.2)?;
    /// #     worksheet.write(3, 0, 37.5)?;
    /// #     worksheet.write(4, 0, 52.1)?;
    /// #     worksheet.write(5, 0, 58.9)?;
    /// #
    ///     // Create a trendline.
    ///     let mut trendline = ChartTrendline::new();
    ///     trendline
    ///         .set_type(ChartTrendlineType::Linear)
    ///         .set_name("My trend name");
    ///
    ///     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with a trendline.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_trendline(&trendline);
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_trendline_set_name.png">
    ///
    pub fn set_name(&mut self, name: impl Into<String>) -> &mut ChartTrendline {
        self.name = name.into();
        self
    }

    /// Set the forward period for a chart trendline.
    ///
    /// Extend the trendline forward by a multiplier of the default length.
    ///
    /// # Parameters
    ///
    /// * `period` - The forward period value.
    ///
    pub fn set_forward_period(&mut self, period: impl Into<f64>) -> &mut ChartTrendline {
        self.forward_period = period.into();
        self
    }

    /// Set the backward period for a chart trendline.
    ///
    /// Extend the trendline backward by a multiplier of the default length.
    ///
    /// # Parameters
    ///
    /// * `period` - The backward period value.
    ///
    pub fn set_backward_period(&mut self, period: impl Into<f64>) -> &mut ChartTrendline {
        self.backward_period = period.into();
        self
    }

    /// Display the trendline equation for a chart trendline.
    ///
    /// Note, the equation is calculated by Excel at runtime. It isn't
    /// calculated by `rust_xlsxwriter` or stored in the Excel file format.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn display_equation(&mut self, enable: bool) -> &mut ChartTrendline {
        self.display_equation = enable;
        self
    }

    /// Display the R-squared value for a chart trendline.
    ///
    /// Display the R-squared [coefficient of determination] for the trendline
    /// as an indicator of how accurate the fit is.
    ///
    /// Note, the R-squared value is calculated by Excel at runtime. It isn't
    /// calculated by `rust_xlsxwriter` or stored in the Excel file format.
    ///
    ///
    /// [coefficient of determination]:
    ///     https://en.wikipedia.org/wiki/Coefficient_of_determination
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn display_r_squared(&mut self, enable: bool) -> &mut ChartTrendline {
        self.display_r_squared = enable;
        self
    }

    /// Set the Y-axis intercept for a chart trendline.
    ///
    /// Set the point where the trendline will intercept the Y-axis.
    ///
    /// # Parameters
    ///
    /// * `intercept` - The intercept with the Y-axis.
    ///
    pub fn set_intercept(&mut self, intercept: impl Into<f64>) -> &mut ChartTrendline {
        self.intercept = Some(intercept.into());
        self
    }

    /// Delete/hide the trendline name from the chart legend.
    ///
    /// The `delete_from_legend()` method deletes/hides the trendline name from
    /// the chart legend. This is often desirable since the trendlines are
    /// generally obvious relative to their series and their names can clutter
    /// the chart legend.
    ///
    /// See also the [`ChartSeries::delete_from_legend()`] and the
    /// [`ChartLegend::delete_entries()`] methods.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// An example of adding a trendline to a chart data series. This
    /// demonstrates deleting/hiding the trendline name from the chart legend.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_trendline_delete_from_legend.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartTrendline, ChartTrendlineType, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 11.1)?;
    /// #     worksheet.write(1, 0, 18.8)?;
    /// #     worksheet.write(2, 0, 33.2)?;
    /// #     worksheet.write(3, 0, 37.5)?;
    /// #     worksheet.write(4, 0, 52.1)?;
    /// #     worksheet.write(5, 0, 58.9)?;
    /// #
    ///     // Create a trendline.
    ///     let mut trendline = ChartTrendline::new();
    ///     trendline
    ///         .set_type(ChartTrendlineType::Linear)
    ///         .delete_from_legend(true);
    ///
    ///     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///
    ///     // Add a data series with a trendline.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_trendline(&trendline);
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_trendline_delete_from_legend.png">
    ///
    /// The default display without deleting the name from the legend would look
    /// like this:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_trendline_delete_from_legend2.png">
    ///
    pub fn delete_from_legend(&mut self, enable: bool) -> &mut ChartTrendline {
        self.delete_from_legend = enable;
        self
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
/// The `ChartTrendlineType` enum defines the trendline types of a
/// [`ChartSeries`].
///
/// The following are the trendline types supported by Excel.
///
/// <img src="https://rustxlsxwriter.github.io/images/trendline_types.png">
///
/// The trendline type is used in conjunction with the
/// [`ChartTrendline::set_type()`] method and a [`ChartSeries`].
///
pub enum ChartTrendlineType {
    /// Don't show any trendline for the data series. The default.
    None,

    /// Display an exponential fit trendline.
    Exponential,

    /// Display a linear best fit trendline.
    Linear,

    /// Display a logarithmic best fit trendline.
    Logarithmic,

    /// Display a polynomial fit trendline. The order of the polynomial can be
    /// specified in the range 2-6.
    Polynomial(u8),

    /// Display a power fit trendline.
    Power,

    /// Display a moving average trendline. The period of the moving average can
    /// be specified in the range 2-4.
    MovingAverage(u8),
}

impl fmt::Display for ChartTrendlineType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Power => write!(f, "power"),
            Self::Linear => write!(f, "linear"),
            Self::Exponential => write!(f, "exp"),
            Self::Logarithmic => write!(f, "log"),
            Self::Polynomial(_) => write!(f, "poly"),
            Self::MovingAverage(_) => write!(f, "movingAvg"),
        }
    }
}

/// The `ChartGradientFill` struct represents a gradient fill for a chart
/// element.
///
/// The [`ChartGradientFill`] struct represents the formatting properties for
/// the gradient fill of a Chart element. In Excel a gradient fill is comprised
/// of two or more colors that are blended gradually along a gradient.
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/gradient_fill_options.png">
///
/// `ChartGradientFill` is a sub property of the [`ChartFormat`] struct and is
/// used with the [`ChartFormat::set_gradient_fill()`] method.
///
/// It is used in conjunction with the [`Chart`] struct.
///
///
/// # Examples
///
/// An example of setting a gradient fill for a chart element.
///
/// ```
/// # // This code is available in examples/doc_chart_gradient_fill.rs
/// #
/// use rust_xlsxwriter::{
///     Chart, ChartGradientFill, ChartGradientStop, ChartType, Workbook, XlsxError,
/// };
///
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///     let worksheet = workbook.add_worksheet();
///
///     // Add some data for the chart.
///     worksheet.write(0, 0, 10)?;
///     worksheet.write(1, 0, 40)?;
///     worksheet.write(2, 0, 50)?;
///     worksheet.write(3, 0, 20)?;
///     worksheet.write(4, 0, 10)?;
///     worksheet.write(5, 0, 50)?;
///
///     // Create a new chart.
///   let mut chart = Chart::new(ChartType::Column);
///
///     // Add a data series with formatting.
///     chart
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_format(ChartGradientFill::new().set_gradient_stops(&[
///             ChartGradientStop::new("#963735", 0),
///             ChartGradientStop::new("#F1DCDB", 100),
///         ]));
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
/// <img src="https://rustxlsxwriter.github.io/images/chart_gradient_fill.png">
///
#[derive(Clone, PartialEq)]
pub struct ChartGradientFill {
    gradient_type: ChartGradientFillType,
    gradient_stops: Vec<ChartGradientStop>,
    angle: u16,
}

impl Default for ChartGradientFill {
    fn default() -> Self {
        Self::new()
    }
}

// -----------------------------------------------------------------------
// ChartGradientFill
// -----------------------------------------------------------------------

impl ChartGradientFill {
    /// Create a new `ChartGradientFill` object to represent a Chart gradient fill.
    ///
    pub fn new() -> ChartGradientFill {
        ChartGradientFill {
            gradient_type: ChartGradientFillType::Linear,
            gradient_stops: vec![],
            angle: 90,
        }
    }

    /// Set the type of the gradient fill.
    ///
    /// Change the default type of the gradient fill to one of the styles
    /// supported by Excel.
    ///
    /// The four gradient types supported by Excel are:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_gradient_fill_types.png">
    ///
    /// # Parameters
    ///
    /// `gradient_type`: a [`ChartGradientFillType`] enum value.
    ///
    /// # Examples
    ///
    /// An example of setting a gradient fill for a chart element with a non-default
    /// gradient type.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_gradient_fill_set_type.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartGradientFill, ChartGradientFillType, ChartGradientStop, ChartType, Workbook,
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
    ///             ChartGradientFill::new()
    ///                 .set_type(ChartGradientFillType::Rectangular)
    ///                 .set_gradient_stops(&[
    ///                     ChartGradientStop::new("#963735", 0),
    ///                     ChartGradientStop::new("#F1DCDB", 100),
    ///                 ]),
    ///         );
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/chart_gradient_fill_set_type.png">
    ///
    pub fn set_type(&mut self, gradient_type: ChartGradientFillType) -> &mut ChartGradientFill {
        self.gradient_type = gradient_type;
        self
    }

    /// Set the gradient stops (data points) for a chart gradient fill.
    ///
    /// A gradient stop, encapsulated by the [`ChartGradientStop`] struct,
    /// represents the properties of a data point that is used to generate a
    /// gradient fill.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/gradient_fill_options.png">
    ///
    /// Excel supports between 2 and 10 gradient stops which define the a color
    /// and its position in the gradient as a percentage. These colors and
    /// positions are used to interpolate a gradient fill.
    ///
    /// # Parameters
    ///
    /// `gradient_stops`: A slice ref of [`ChartGradientStop`] values. As in
    /// Excel there must be between 2 and 10 valid gradient stops.
    ///
    /// # Examples
    ///
    /// An example of setting a gradient fill for a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_gradient_stops.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartGradientFill, ChartGradientStop, ChartType, Workbook, XlsxError,
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
    ///     // Set the properties of the gradient stops.
    ///     let gradient_stops = [
    ///         ChartGradientStop::new("#156B13", 0),
    ///         ChartGradientStop::new("#9CB86E", 50),
    ///         ChartGradientStop::new("#DDEBCF", 100),
    ///     ];
    ///
    ///     // Add a data series with formatting.
    ///     chart
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(ChartGradientFill::new().set_gradient_stops(&gradient_stops));
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
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
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_gradient_stops.png">
    ///
    /// Note, it can be clearer to add the gradient stops directly to the format
    /// as follows. This gives the same output as above.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_gradient_stops2.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Chart, ChartGradientFill, ChartGradientStop, ChartType, Workbook, XlsxError,
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
    ///         .set_format(ChartGradientFill::new().set_gradient_stops(&[
    ///             ChartGradientStop::new("#156B13", 0),
    ///             ChartGradientStop::new("#9CB86E", 50),
    ///             ChartGradientStop::new("#DDEBCF", 100),
    ///         ]));
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    pub fn set_gradient_stops(
        &mut self,
        gradient_stops: &[ChartGradientStop],
    ) -> &mut ChartGradientFill {
        let mut valid_gradient_stops = vec![];

        for gradient_stop in gradient_stops {
            if gradient_stop.is_valid() {
                valid_gradient_stops.push(gradient_stop.clone());
            }
        }

        if (2..=10).contains(&valid_gradient_stops.len()) {
            self.gradient_stops = valid_gradient_stops;
        } else {
            eprintln!("Gradient stops must contain between 2 and 10 valid entries.");
        }

        self
    }

    /// Set the angle of the linear gradient fill type.
    ///
    /// # Parameters
    ///
    /// * `angle` - The angle of the linear gradient fill in the range `0 <=
    ///   angle < 360`. The default angle is 90 degrees.
    ///
    pub fn set_angle(&mut self, angle: u16) -> &mut ChartGradientFill {
        if (0..360).contains(&angle) {
            self.angle = angle;
        } else {
            eprintln!("Gradient angle '{angle}' must be in the Excel range 0 <= angle < 360");
        }
        self
    }
}

/// The `ChartGradientStop` struct represents a gradient fill data point.
///
/// The [`ChartGradientStop`] struct represents the properties of a data point
/// (a stop) that is used to generate a gradient fill.
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/gradient_fill_options.png">
///
/// Excel supports between 2 and 10 gradient stops which define the a color and
/// its position in the gradient as a percentage. These colors and positions
/// are used to interpolate a gradient fill.
///
/// Gradient formats are generally used with the
/// [`ChartGradientFill::set_gradient_stops()`] method and
/// [`ChartGradientFill`].
///
/// # Examples
///
/// An example of setting a gradient fill for a chart element.
///
/// ```
/// # // This code is available in examples/doc_chart_gradient_stops.rs
/// #
/// # use rust_xlsxwriter::{
/// #     Chart, ChartGradientFill, ChartGradientStop, ChartType, Workbook, XlsxError,
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
///     // Set the properties of the gradient stops.
///     let gradient_stops = [
///         ChartGradientStop::new("#156B13", 0),
///         ChartGradientStop::new("#9CB86E", 50),
///         ChartGradientStop::new("#DDEBCF", 100),
///     ];
///
///     // Add a data series with formatting.
///     chart
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_format(ChartGradientFill::new().set_gradient_stops(&gradient_stops));
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
/// <img src="https://rustxlsxwriter.github.io/images/chart_gradient_stops.png">
///
///
#[derive(Clone, PartialEq)]
pub struct ChartGradientStop {
    color: Color,
    position: u8,
}

impl ChartGradientStop {
    /// Create a new `ChartGradientStop` object to represent a Chart gradient fill stop.
    ///
    /// # Parameters
    ///
    /// * `color` - The gradient stop color property defined by a [`Color`] enum
    ///   value.
    /// * `position` - The gradient stop position in the range 0-100.
    ///
    /// # Examples
    ///
    /// An example of creating gradient stops for a gradient fill for a chart element.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_gradient_stops_new.rs
    /// #
    /// # use rust_xlsxwriter::ChartGradientStop;
    /// #
    /// # #[allow(unused_variables)]
    /// # fn main() {
    ///     let gradient_stops = [
    ///         ChartGradientStop::new("#156B13", 0),
    ///         ChartGradientStop::new("#9CB86E", 50),
    ///         ChartGradientStop::new("#DDEBCF", 100),
    ///     ];
    /// # }
    /// ```
    pub fn new<T>(color: T, position: u8) -> ChartGradientStop
    where
        T: IntoColor,
    {
        let color = color.new_color();

        // Check and warn but don't raise error since this is too deeply nested.
        // It will be rechecked and rejected at use.
        if !color.is_valid() {
            eprintln!("Gradient stop color isn't valid.");
        }
        if !(0..=100).contains(&position) {
            eprintln!("Gradient stop '{position}' outside Excel range: 0 <= position <= 100.");
        }

        ChartGradientStop { color, position }
    }

    // Check for valid gradient stop properties.
    pub(crate) fn is_valid(&self) -> bool {
        self.color.is_valid() && (0..=100).contains(&self.position)
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
/// The `ChartGradientFillType` enum defines the gradient types of a
/// [`ChartGradientFill`].
///
/// The four gradient types supported by Excel are:
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_gradient_fill_types.png">
///
pub enum ChartGradientFillType {
    /// The gradient runs linearly from the top of the area vertically to the
    /// bottom. This is the default.
    Linear,

    /// The gradient runs radially from the bottom right of the area vertically
    /// to the top left.
    Radial,

    /// The gradient runs in a rectangular pattern from the bottom right of the
    /// area vertically to the top left.
    Rectangular,

    /// The gradient runs in a rectangular pattern from the center of the area
    /// to the outer vertices.
    Path,
}

// -----------------------------------------------------------------------
// ChartErrorBars
// -----------------------------------------------------------------------

/// The `ChartErrorBars` struct represents the error bars for a chart series.
///
/// Error bars on Excel charts allow you to show margins of error for a series
/// based on measures such as Standard Deviation, Standard Error, Fixed values,
/// Percentages or even custom defined ranges.
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/chart_error_bars_options.png">
///
/// The `ChartErrorBars` struct can be added to a series via the
/// [`ChartSeries::set_y_error_bars()`] and [`ChartSeries::set_x_error_bars()`]
/// methods.
///
/// # Examples
///
/// An example of adding error bars to a chart data series.
///
/// ```
/// # // This code is available in examples/doc_chart_error_bars_intro.rs
/// #
/// # use rust_xlsxwriter::{
/// #     Chart, ChartErrorBars, ChartErrorBarsType, ChartLine, ChartType, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the chart.
/// #     worksheet.write(0, 0, 11.1)?;
/// #     worksheet.write(1, 0, 18.8)?;
/// #     worksheet.write(2, 0, 33.2)?;
/// #     worksheet.write(3, 0, 37.5)?;
/// #     worksheet.write(4, 0, 52.1)?;
/// #     worksheet.write(5, 0, 58.9)?;
/// #
/// #     // Create a new chart.
///     let mut chart = Chart::new(ChartType::Line);
///
///     // Add a data series with error bars.
///     chart
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_y_error_bars(
///             ChartErrorBars::new()
///                 .set_type(ChartErrorBarsType::StandardError)
///                 .set_format(ChartLine::new().set_color("#FF0000")),
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
/// src="https://rustxlsxwriter.github.io/images/chart_error_bars_intro.png">
///
#[derive(Clone, PartialEq)]
pub struct ChartErrorBars {
    has_end_cap: bool,
    error_type: ChartErrorBarsType,
    direction: ChartErrorBarsDirection,
    format: ChartFormat,
    pub(crate) plus_range: ChartRange,
    pub(crate) minus_range: ChartRange,
}

impl Default for ChartErrorBars {
    fn default() -> Self {
        Self::new()
    }
}

impl ChartErrorBars {
    /// Create a new `ChartErrorBars` object to represent Chart series error bars.
    ///
    pub fn new() -> ChartErrorBars {
        ChartErrorBars {
            has_end_cap: true,
            error_type: ChartErrorBarsType::StandardError,
            direction: ChartErrorBarsDirection::Both,
            format: ChartFormat::default(),
            plus_range: ChartRange::default(),
            minus_range: ChartRange::default(),
        }
    }

    /// Set the type of the Chart series error bars.
    ///
    /// Set the error bar type to one of the Excel allowable amounts represented
    /// by the [`ChartErrorBarsType`] enum.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/chart_error_bars_types.png">
    ///
    /// # Parameters
    ///
    /// * `error_type` - A [`ChartErrorBarsType`] enum reference.
    ///
    pub fn set_type(&mut self, error_type: ChartErrorBarsType) -> &mut ChartErrorBars {
        match &error_type {
            ChartErrorBarsType::FixedValue(value) => {
                if *value <= 0.0 {
                    eprintln!("Error bar Fixed Value '{value}' must be > 0.0 in Excel");
                    return self;
                }
            }
            ChartErrorBarsType::Percentage(value) => {
                if *value < 0.0 {
                    eprintln!("Error bar Percentage '{value}' must be >= 0.0 in Excel");
                    return self;
                }
            }
            ChartErrorBarsType::StandardDeviation(value) => {
                if *value < 0.0 {
                    eprintln!("Error bar Standard Deviation '{value}' must be >= 0.0 in Excel");
                    return self;
                }
            }
            ChartErrorBarsType::Custom(plus, minus) => {
                self.plus_range = (*plus).clone();
                self.minus_range = (*minus).clone();
            }
            ChartErrorBarsType::StandardError => {}
        }

        self.error_type = error_type;

        self
    }

    /// Set the direction of a Chart series error bars.
    ///
    /// The [`ChartErrorBarsDirection`] enum defines the error bar direction for a
    /// chart series.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_error_bars_directions.png">
    ///
    /// # Parameters
    ///
    /// * `direction` - A [`ChartErrorBarsDirection`] enum reference.
    ///
    pub fn set_direction(&mut self, direction: ChartErrorBarsDirection) -> &mut ChartErrorBars {
        self.direction = direction;
        self
    }

    /// Set the end cap on/off for a Chart series error bars.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is on by default.
    ///
    pub fn set_end_cap(&mut self, enable: bool) -> &mut ChartErrorBars {
        self.has_end_cap = enable;
        self
    }

    /// Set the formatting properties for a chart series error bars.
    ///
    /// Set the formatting properties for a chart series via a [`ChartFormat`]
    /// object or a sub struct that implements [`IntoChartFormat`].
    ///
    /// For error bars the only formatting supported by Excel is
    /// [`ChartFormat::set_line()`].
    ///
    /// # Parameters
    ///
    /// `format`: A [`ChartFormat`] struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// [`IntoChartFormat`] for details.
    ///
    pub fn set_format<T>(&mut self, format: T) -> &mut ChartErrorBars
    where
        T: IntoChartFormat,
    {
        self.format = format.new_chart_format();
        self
    }
}

#[derive(Clone, PartialEq)]
/// The `ChartErrorBarsType` enum defines the type of a chart series
/// [`ChartErrorBars`].
///
/// The following enum values represent the error bar types that are available
/// in Excel.
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/chart_error_bars_types.png">
///
pub enum ChartErrorBarsType {
    /// Set a fixed value for the positive and negative error bars. In Excel
    /// this must be > 0.0.
    FixedValue(f64),

    /// Set a percentage for the positive and negative error bars. In Excel this
    /// must be >= 0.0.
    Percentage(f64),

    /// Set a multiple of the standard deviation for the positive and negative
    /// error bars. In Excel this must be >= 0.0.
    StandardDeviation(f64),

    /// Set a the standard error value for the positive and negative error bars.
    /// This is the default.
    StandardError,

    /// Set custom values for the error bars based on a range of worksheet
    /// values like `Sheet1!$B$1:$B$3` (single value) or `Sheet1!$B$1:$B$5` (a
    /// range to match the number of point in the series). Single values are
    /// repeated for each point in the chart, like `FixedValue`. The `plus` and
    /// `minus` values must be set separately using [`ChartRange`] instances.
    Custom(ChartRange, ChartRange),
}

impl fmt::Display for ChartErrorBarsType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Custom(_, _) => write!(f, "cust"),
            Self::StandardError => write!(f, "stdErr"),
            Self::FixedValue(_) => write!(f, "fixedVal"),
            Self::Percentage(_) => write!(f, "percentage"),
            Self::StandardDeviation(_) => write!(f, "stdDev"),
        }
    }
}

#[derive(Clone, Copy, PartialEq)]
/// The `ChartErrorBarsDirection` enum defines the error bar direction for a
/// chart series [`ChartErrorBars`].
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_error_bars_directions.png">
///
pub enum ChartErrorBarsDirection {
    /// The error bars extend in both directions. This is the default.
    Both,

    /// The error bars extend in the negative direction only.
    Minus,

    /// The error bars extend in the positive direction only.
    Plus,
}

impl fmt::Display for ChartErrorBarsDirection {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Both => write!(f, "both"),
            Self::Minus => write!(f, "minus"),
            Self::Plus => write!(f, "plus"),
        }
    }
}

// -----------------------------------------------------------------------
// ChartDataTable
// -----------------------------------------------------------------------

/// The `ChartDataTable` struct represents an optional data table displayed
/// below the chart.
///
/// A chart data table in Excel is an additional table below a chart that shows
/// the plotted data in tabular form.
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_data_table.png">
///
/// The chart data table has the following default properties which can be set
/// with the methods outlined below.
///
/// The `ChartDataTable` struct is used in conjunction with the
/// [`Chart::set_data_table()`] method.
///
///  <img src="https://rustxlsxwriter.github.io/images/chart_data_table_options.png">
///
/// # Examples
///
/// An example of adding a data table to a chart.
///
/// ```
/// # // This code is available in examples/doc_chart_set_data_table.rs
/// #
/// # use rust_xlsxwriter::{Chart, ChartDataTable, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the chart.
/// #     let data = [[1, 2, 3], [2, 4, 6], [3, 6, 9], [4, 8, 12], [5, 10, 15]];
/// #     for (row_num, row_data) in data.iter().enumerate() {
/// #         for (col_num, col_data) in row_data.iter().enumerate() {
/// #             worksheet.write_number(row_num as u32, col_num as u16, *col_data)?;
/// #         }
/// #     }
/// #
/// #     // Create a new chart.
///     let mut chart = Chart::new_column();
///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
///     chart.add_series().set_values("Sheet1!$B$1:$B$5");
///     chart.add_series().set_values("Sheet1!$C$1:$C$5");
///
///     // Add a default data table to the chart.
///     let table = ChartDataTable::default();
///     chart.set_data_table(&table);
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
/// <img src="https://rustxlsxwriter.github.io/images/chart_set_data_table.png">
///
#[derive(Clone, PartialEq)]
pub struct ChartDataTable {
    show_horizontal_borders: bool,
    show_vertical_borders: bool,
    show_outline_borders: bool,
    show_legend_keys: bool,
    font: Option<ChartFont>,
    format: ChartFormat,
}

impl Default for ChartDataTable {
    fn default() -> Self {
        Self::new()
    }
}

impl ChartDataTable {
    /// Create a new `ChartDataTable` object to represent a Chart Data Table.
    ///
    pub fn new() -> ChartDataTable {
        ChartDataTable {
            show_horizontal_borders: true,
            show_vertical_borders: true,
            show_outline_borders: true,
            show_legend_keys: false,
            font: None,
            format: ChartFormat::default(),
        }
    }

    /// Turn on/off the horizontal border lines for a chart data table.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is on by default.
    ///
    pub fn show_horizontal_borders(mut self, enable: bool) -> ChartDataTable {
        self.show_horizontal_borders = enable;
        self
    }

    /// Turn on/off the vertical border lines for a chart data table.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is on by default.
    ///
    pub fn show_vertical_borders(mut self, enable: bool) -> ChartDataTable {
        self.show_vertical_borders = enable;
        self
    }

    /// Turn on/off the outline border lines for a chart data table.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is on by default.
    ///
    pub fn show_outline_borders(mut self, enable: bool) -> ChartDataTable {
        self.show_outline_borders = enable;
        self
    }

    /// Turn on/off the legend keys for a chart data table.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn show_legend_keys(mut self, enable: bool) -> ChartDataTable {
        self.show_legend_keys = enable;
        self
    }

    /// Set the formatting properties for a chart data table.
    ///
    /// Set the formatting properties for a chart data table via a [`ChartFormat`]
    /// object or a sub struct that implements [`IntoChartFormat`].
    ///
    /// The formatting that can be applied via a [`ChartFormat`] object are:
    ///
    /// - [`ChartFormat::set_solid_fill()`]: Set the [`ChartSolidFill`] properties.
    /// - [`ChartFormat::set_pattern_fill()`]: Set the [`ChartPatternFill`] properties.
    /// - [`ChartFormat::set_gradient_fill()`]: Set the [`ChartGradientFill`] properties.
    /// - [`ChartFormat::set_no_fill()`]: Turn off the fill for the chart object.
    /// - [`ChartFormat::set_line()`]: Set the [`ChartLine`] properties.
    /// - [`ChartFormat::set_border()`]: Set the [`ChartBorder`] properties.
    ///   A synonym for [`ChartLine`] depending on context.
    /// - [`ChartFormat::set_no_line()`]: Turn off the line for the chart object.
    /// - [`ChartFormat::set_no_border()`]: Turn off the border for the chart object.
    ///
    /// # Parameters
    ///
    /// `format`: A [`ChartFormat`] struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// [`IntoChartFormat`] for details.
    ///
    pub fn set_format<T>(mut self, format: T) -> ChartDataTable
    where
        T: IntoChartFormat,
    {
        self.format = format.new_chart_format();
        self
    }

    /// Set the font properties of a chart data table.
    ///
    /// Set the font properties of a chart data table using a [`ChartFont`]
    /// reference. Example font properties that can be set are:
    ///
    /// - [`ChartFont::set_bold()`]
    /// - [`ChartFont::set_italic()`]
    /// - [`ChartFont::set_name()`]
    /// - [`ChartFont::set_size()`]
    /// - [`ChartFont::set_rotation()`]
    ///
    /// See [`ChartFont`] for full details.
    ///
    /// # Parameters
    ///
    /// `font`: A [`ChartFont`] struct reference to represent the font
    /// properties.
    ///

    ///
    pub fn set_font(mut self, font: &ChartFont) -> ChartDataTable {
        self.font = Some(font.clone());
        self
    }
}

// -----------------------------------------------------------------------
// ChartAxisCrossing
// -----------------------------------------------------------------------

/// The `ChartAxisCrossing` enum defines the [`ChartAxis`] crossing point for
/// the opposite axis.
///
/// By default Excel sets chart axes to cross at 0. If required you can use
/// [`ChartAxis::set_crossing()`] and [`ChartAxisCrossing`] to define another
/// point where the opposite axis will cross the current axis.
///
/// # Examples
///
/// A chart example demonstrating setting the point where the axes will cross.
///
/// ```
/// # // This code is available in examples/doc_chart_axis_set_crossing.rs
/// #
/// # use rust_xlsxwriter::{Chart, ChartAxisCrossing, ChartType, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the chart.
/// #     worksheet.write(0, 0, "North")?;
/// #     worksheet.write(1, 0, "South")?;
/// #     worksheet.write(2, 0, "East")?;
/// #     worksheet.write(3, 0, "West")?;
/// #     worksheet.write(0, 1, 10)?;
/// #     worksheet.write(1, 1, 35)?;
/// #     worksheet.write(2, 1, 40)?;
/// #     worksheet.write(3, 1, 25)?;
/// #
/// #     // Create a new chart.
///     let mut chart = Chart::new(ChartType::Column);
///
///     // Add a data series using Excel formula syntax to describe the range.
///     chart
///         .add_series()
///         .set_categories("Sheet1!$A$1:$A$5")
///         .set_values("Sheet1!$B$1:$B$5");
///
///     // Set the X-axis crossing at a category index.
///     chart
///         .x_axis()
///         .set_crossing(ChartAxisCrossing::CategoryNumber(3));
///
///     // Set the Y-axis crossing at a value.
///     chart
///         .y_axis()
///         .set_crossing(ChartAxisCrossing::AxisValue(20.0));
///
///     // Hide legend for clarity.
///     chart.legend().set_hidden();
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
/// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_crossing1.png">
///
/// For reference here is the default chart without default crossings:
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_axis_set_crossing2.png">
///
///
#[derive(Clone, Copy, PartialEq)]
pub enum ChartAxisCrossing {
    /// The axis crossing is at the default value which is generally zero. This
    /// is the default.
    Automatic,

    /// The axis crossing is at the minimum value for the axis.
    Min,

    /// The axis crossing is at the maximum value for the axis.
    Max,

    /// The axis crossing is at a category index number.
    ///
    /// This is for Category style axes only. For example say you are plotting 4
    /// categories on the X-axis ("North", "South", "East", "West"). By setting
    /// the category number as `CategoryNumber(3)` the Y-axis will cross at
    /// "East".
    ///
    /// See [Chart Value and Category
    /// Axes](crate::chart#chart-value-and-category-axes) for an
    /// explanation of the difference between Value and Category axes in Excel.
    ///
    CategoryNumber(u32),

    /// The axis crossing is at a value.
    ///
    /// This is for Value and Date style axes only.
    ///
    /// See [Chart Value and Category
    /// Axes](crate::chart#chart-value-and-category-axes) for an
    /// explanation of the difference between Value and Category axes in Excel.
    ///
    AxisValue(f64),
}

impl fmt::Display for ChartAxisCrossing {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Min => write!(f, "min"),
            Self::Max => write!(f, "max"),
            Self::Automatic => write!(f, "autoZero"),
            Self::AxisValue(value) => write!(f, "{value}"),
            Self::CategoryNumber(index) => write!(f, "{index}"),
        }
    }
}

// -----------------------------------------------------------------------
// ChartAxisLabelAlignment
// -----------------------------------------------------------------------

/// The `ChartAxisLabelAlignment` enum defines the [`ChartAxis`] crossing point for
/// the opposite axis.
///
#[derive(Clone, Copy, PartialEq)]
pub enum ChartAxisLabelAlignment {
    /// Center the axis label with the tick mark. This is the default.
    Center,

    /// Set the axis label to the left of the tick mark.
    Left,

    /// Set the axis label to the right of the tick mark.
    Right,
}

impl fmt::Display for ChartAxisLabelAlignment {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Left => write!(f, "l"),
            Self::Right => write!(f, "r"),
            Self::Center => write!(f, "ctr"),
        }
    }
}
