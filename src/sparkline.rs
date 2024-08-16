// Sparkline - A module to represent an Excel sparkline.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! # Working with Sparklines
//!
//! Sparklines are a feature of Excel 2010+ which allows you to add small charts
//! to worksheet cells. These are useful for showing data trends in a compact
//! visual format.
//!
//! <img src="https://rustxlsxwriter.github.io/images/sparklines1.png">
//!
//! The following code was used to generate the above file.
//!
//!
//! ```rust
//! # // This code is available in examples/app_sparklines1.rs
//! #
//! use rust_xlsxwriter::{Sparkline, SparklineType, Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     // Create a new Excel file object.
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Some sample data to plot.
//!     let data = [[-2, 2, 3, -1, 0], [30, 20, 33, 20, 15], [1, -1, -1, 1, -1]];
//!
//!     worksheet.write_row_matrix(0, 0, data)?;
//!
//!     // Add a line sparkline (the default) with markers.
//!     let sparkline1 = Sparkline::new()
//!         .set_range(("Sheet1", 0, 0, 0, 4))
//!         .show_markers(true);
//!
//!     worksheet.add_sparkline(0, 5, &sparkline1)?;
//!
//!     // Add a column sparkline with non-default style.
//!     let sparkline2 = Sparkline::new()
//!         .set_range(("Sheet1", 1, 0, 1, 4))
//!         .set_type(SparklineType::Column)
//!         .set_style(12);
//!
//!     worksheet.add_sparkline(1, 5, &sparkline2)?;
//!
//!     // Add a win/loss sparkline with negative values highlighted.
//!     let sparkline3 = Sparkline::new()
//!         .set_range(("Sheet1", 2, 0, 2, 4))
//!         .set_type(SparklineType::WinLose)
//!         .show_negative_points(true);
//!
//!     worksheet.add_sparkline(2, 5, &sparkline3)?;
//!
//!     // Save the file to disk.
//!     workbook.save("sparklines1.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! In Excel sparklines can be added as a single entity in a cell that refers to
//! a 1D data range or as a "group" sparkline that is applied across a 1D range
//! and refers to data in a 2D range. A grouped sparkline uses one sparkline for
//! the specified range and any changes to it are applied to the entire
//! sparkline group.
//!
//! The [`Worksheet::add_sparkline()`](crate::Worksheet::add_sparkline) method
//! shown allows you to add a sparkline to a single cell that displays data from
//! a 1D range of cells whereas the
//! [`Worksheet::add_sparkline_group()`](crate::Worksheet::add_sparkline_group())
//! method applies the group sparkline to a range.
//!
//! Both of these methods are shown in the example below along with a
//! demonstration of various properties that can be set for Excel sparklines:
//!
//! <img src="https://rustxlsxwriter.github.io/images/sparklines2.png">
//!
//! The following code was used to generate the above file.
//!
//! ```rust
//! # // This code is available in examples/app_sparklines2.rs
//! #
//! use rust_xlsxwriter::{Format, Sparkline, SparklineType, Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     // Create a new Excel file object.
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet1 = workbook.add_worksheet();
//!     let mut row = 1;
//!
//!     // Set the columns widths to make the output clearer.
//!     worksheet1.set_column_width(0, 14)?;
//!     worksheet1.set_column_width(1, 50)?;
//!     worksheet1.set_zoom(150);
//!
//!     // Add some headings.
//!     let bold = Format::new().set_bold();
//!     worksheet1.write_with_format(0, 0, "Sparkline", &bold)?;
//!     worksheet1.write_with_format(0, 1, "Description", &bold)?;
//!
//!     //
//!     // Add a default line sparkline.
//!     //
//!     let text = "A default line sparkline.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new().set_range(("Sheet2", 0, 0, 0, 9));
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Add a default column sparkline.
//!     //
//!     let text = "A default column sparkline.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 1, 0, 1, 9))
//!         .set_type(SparklineType::Column);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Add a default win/loss sparkline.
//!     //
//!     let text = "A default win/loss sparkline.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 2, 0, 2, 9))
//!         .set_type(SparklineType::WinLose);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 2;
//!
//!     //
//!     // Add a line sparkline with markers.
//!     //
//!     let text = "Line with markers.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 0, 0, 0, 9))
//!         .show_markers(true);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Add a line sparkline with high and low points.
//!     //
//!     let text = "Line with high and low points.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 0, 0, 0, 9))
//!         .show_high_point(true)
//!         .show_low_point(true);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Add a line sparkline with first and last points.
//!     //
//!     let text = "Line with first and last point markers.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 0, 0, 0, 9))
//!         .show_first_point(true)
//!         .show_last_point(true);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Add a line sparkline with negative point markers.
//!     //
//!     let text = "Line with negative point markers.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 0, 0, 0, 9))
//!         .show_negative_points(true);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Add a line sparkline with axis.
//!     //
//!     let text = "Line with axis.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 0, 0, 0, 9))
//!         .show_axis(true);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 2;
//!
//!     //
//!     // Add a column sparkline with style 1. The default style.
//!     //
//!     let text = "Column with style 1. The default.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 1, 0, 1, 9))
//!         .set_type(SparklineType::Column)
//!         .set_style(1);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Add a column sparkline with style 2.
//!     //
//!     let text = "Column with style 2.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 1, 0, 1, 9))
//!         .set_type(SparklineType::Column)
//!         .set_style(2);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Add a column sparkline with style 3.
//!     //
//!     let text = "Column with style 3.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 1, 0, 1, 9))
//!         .set_type(SparklineType::Column)
//!         .set_style(3);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Add a column sparkline with style 4.
//!     //
//!     let text = "Column with style 4.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 1, 0, 1, 9))
//!         .set_type(SparklineType::Column)
//!         .set_style(4);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Add a column sparkline with style 5.
//!     //
//!     let text = "Column with style 5.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 1, 0, 1, 9))
//!         .set_type(SparklineType::Column)
//!         .set_style(5);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Add a column sparkline with style 6.
//!     //
//!     let text = "Column with style 6.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 1, 0, 1, 9))
//!         .set_type(SparklineType::Column)
//!         .set_style(6);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Add a column sparkline with a user defined color.
//!     //
//!     let text = "Column with a user defined color.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 1, 0, 1, 9))
//!         .set_type(SparklineType::Column)
//!         .set_sparkline_color("#E965E0");
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 2;
//!
//!     //
//!     // Add a win/loss sparkline.
//!     //
//!     let text = "A win/loss sparkline.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 2, 0, 2, 9))
//!         .set_type(SparklineType::WinLose);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Add a win/loss sparkline with negative points highlighted.
//!     //
//!     let text = "A win/loss sparkline with negative points highlighted.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 2, 0, 2, 9))
//!         .set_type(SparklineType::WinLose)
//!         .show_negative_points(true);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 2;
//!
//!     //
//!     // Add a left to right (the default) sparkline.
//!     //
//!     let text = "A left to right column (the default).";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 3, 0, 3, 9))
//!         .set_type(SparklineType::Column)
//!         .set_style(20);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Add a right to left sparkline.
//!     //
//!     let text = "A right to left column.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 3, 0, 3, 9))
//!         .set_type(SparklineType::Column)
//!         .set_style(20)
//!         .set_right_to_left(true);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     row += 1;
//!
//!     //
//!     // Sparkline and text in one cell. This just requires writing text to the
//!     // same cell as the sparkline.
//!     //
//!     let text = "Sparkline and text in one cell.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 3, 0, 3, 9))
//!         .set_type(SparklineType::Column)
//!         .set_style(20);
//!
//!     worksheet1.add_sparkline(row, 0, &sparkline)?;
//!     worksheet1.write(row, 0, "Growth")?;
//!     row += 2;
//!
//!     //
//!     // "A grouped sparkline. User changes are applied to all three. Not that the
//!     // sparkline range is a 2D range and the sparkline is positioned in a 1D
//!     // range of cells.
//!     //
//!     let text = "A grouped sparkline. Changes are applied to all three.";
//!     worksheet1.write(row, 1, text)?;
//!
//!     let sparkline = Sparkline::new()
//!         .set_range(("Sheet2", 4, 0, 6, 9))
//!         .show_markers(true);
//!
//!     worksheet1.add_sparkline_group(row, 0, row + 2, 0, &sparkline)?;
//!
//!     //
//!     // Add a worksheet with the data to plot on a separate worksheet.
//!     //
//!     let worksheet2 = workbook.add_worksheet();
//!
//!     // Some sample data to plot.
//!     let data = [
//!         // Simple line data.
//!         [-2, 2, 3, -1, 0, -2, 3, 2, 1, 0],
//!         // Simple column data.
//!         [30, 20, 33, 20, 15, 5, 5, 15, 10, 15],
//!         // Simple win/loss data.
//!         [1, 1, -1, -1, 1, -1, 1, 1, 1, -1],
//!         // Unbalanced histogram.
//!         [5, 6, 7, 10, 15, 20, 30, 50, 70, 100],
//!         // Data for the grouped sparkline example.
//!         [-2, 2, 3, -1, 0, -2, 3, 2, 1, 0],
//!         [3, -1, 0, -2, 3, 2, 1, 0, 2, 1],
//!         [0, -2, 3, 2, 1, 0, 1, 2, 3, 1],
//!     ];
//!
//!     worksheet2.write_row_matrix(0, 0, data)?;
//!
//!     // Save the file to disk.
//!     workbook.save("sparklines2.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
#![warn(missing_docs)]

use crate::{utility, ChartEmptyCells, ChartRange, ColNum, Color, IntoChartRange, RowNum};

mod tests;

/// The `Sparkline` struct is used to create an object to represent a sparkline
/// that can be inserted into a worksheet.
///
/// Sparklines are a feature of Excel 2010+ which allows you to add small charts
/// to worksheet cells. These are useful for showing data trends in a compact
/// visual format.
///
/// <img src="https://rustxlsxwriter.github.io/images/sparklines1.png">
///
/// The following example was used to generate the above file.
///
/// ```
/// # // This code is available in examples/app_sparklines1.rs
/// #
/// use rust_xlsxwriter::{Sparkline, SparklineType, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Some sample data to plot.
///     let data = [[-2, 2, 3, -1, 0], [30, 20, 33, 20, 15], [1, -1, -1, 1, -1]];
///
///     worksheet.write_row_matrix(0, 0, data)?;
///
///     // Add a line sparkline (the default) with markers.
///     let sparkline1 = Sparkline::new()
///         .set_range(("Sheet1", 0, 0, 0, 4))
///         .show_markers(true);
///
///     worksheet.add_sparkline(0, 5, &sparkline1)?;
///
///     // Add a column sparkline with non-default style.
///     let sparkline2 = Sparkline::new()
///         .set_range(("Sheet1", 1, 0, 1, 4))
///         .set_type(SparklineType::Column)
///         .set_style(12);
///
///     worksheet.add_sparkline(1, 5, &sparkline2)?;
///
///     // Add a win/loss sparkline with negative values highlighted.
///     let sparkline3 = Sparkline::new()
///         .set_range(("Sheet1", 2, 0, 2, 4))
///         .set_type(SparklineType::WinLose)
///         .show_negative_points(true);
///
///     worksheet.add_sparkline(2, 5, &sparkline3)?;
///
///     // Save the file to disk.
///     workbook.save("sparklines1.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// In Excel sparklines can be added as a single entity in a cell that refers to
/// a 1D data range or as a "group" sparkline that is applied across a 1D range
/// and refers to data in a 2D range. A grouped sparkline uses one sparkline for
/// the specified range and any changes to it are applied to the entire
/// sparkline group.
///
/// The [`Worksheet::add_sparkline()`](crate::Worksheet::add_sparkline) method
/// shown allows you to add a sparkline to a single cell that displays data from
/// a 1D range of cells whereas the
/// [`Worksheet::add_sparkline_group()`](crate::Worksheet::add_sparkline_group())
/// method applies the group sparkline to a range.
///
#[derive(Clone)]
pub struct Sparkline {
    pub(crate) series_color: Color,
    pub(crate) negative_points_color: Color,
    pub(crate) axis_color: Color,
    pub(crate) markers_color: Color,
    pub(crate) first_point_color: Color,
    pub(crate) last_point_color: Color,
    pub(crate) high_point_color: Color,
    pub(crate) low_point_color: Color,
    pub(crate) data_range: ChartRange,
    pub(crate) date_range: ChartRange,
    pub(crate) ranges: Vec<(String, String)>,
    pub(crate) sparkline_type: SparklineType,
    pub(crate) show_high_point: bool,
    pub(crate) show_low_point: bool,
    pub(crate) show_first_point: bool,
    pub(crate) show_last_point: bool,
    pub(crate) show_negative_points: bool,
    pub(crate) show_markers: bool,
    pub(crate) show_axis: bool,
    pub(crate) show_hidden_data: bool,
    pub(crate) show_right_to_left: bool,
    pub(crate) show_empty_cells_as: ChartEmptyCells,
    pub(crate) line_weight: Option<f64>,
    pub(crate) custom_min: Option<f64>,
    pub(crate) custom_max: Option<f64>,
    pub(crate) group_max: bool,
    pub(crate) group_min: bool,
    data_row_order: bool,
}

#[allow(clippy::new_without_default)]
impl Sparkline {
    /// Create a new Sparkline struct.
    pub fn new() -> Sparkline {
        Sparkline {
            series_color: Color::Theme(4, 5),
            negative_points_color: Color::Theme(5, 0),
            axis_color: Color::Black,
            markers_color: Color::Theme(4, 5),
            first_point_color: Color::Theme(4, 3),
            last_point_color: Color::Theme(4, 3),
            high_point_color: Color::Theme(4, 0),
            low_point_color: Color::Theme(4, 0),
            data_range: ChartRange::default(),
            date_range: ChartRange::default(),
            ranges: vec![],
            sparkline_type: SparklineType::Line,
            show_high_point: false,
            show_low_point: false,
            show_first_point: false,
            show_last_point: false,
            show_negative_points: false,
            show_markers: false,
            show_axis: false,
            show_hidden_data: false,
            show_right_to_left: false,
            show_empty_cells_as: ChartEmptyCells::Gaps,
            line_weight: None,
            custom_min: None,
            custom_max: None,
            group_max: false,
            group_min: false,
            data_row_order: true,
        }
    }

    /// Set the range of the sparkline data.
    ///
    /// This method is used to set the location of the data from which the
    /// sparkline will be plotted. This constitutes the Y values of the
    /// sparkline.
    ///
    /// By default the X values of the sparkline are taken as evenly spaced
    /// increments. However, it is also possible to specify date values for the
    /// X axis using the
    /// [`Sparkline::set_date_range`](Sparkline::set_date_range) method.
    ///
    /// The range can either be a 1D range when used with
    /// [`Worksheet::add_sparkline()`](crate::Worksheet::add_sparkline()) or a
    /// 2D range with used with
    /// [`Worksheet::add_sparkline_group()`](crate::Worksheet::add_sparkline_group()).
    ///
    /// # Parameters
    ///
    /// - `range`: A 1D or 2D range that contains the data that will be plotted
    ///   in the sparkline. This can specified in different ways, see
    ///   [`IntoChartRange`] for details.
    ///
    pub fn set_range<T>(mut self, range: T) -> Sparkline
    where
        T: IntoChartRange,
    {
        self.data_range = range.new_chart_range();
        self
    }

    /// Set the type of sparkline.
    ///
    /// The type of the sparkline can be one of the following:
    ///
    /// - [`SparklineType::Line`]: A line style sparkline. This is the default.
    ///
    ///   <img src="https://rustxlsxwriter.github.io/images/sparkline_type_line.png">
    ///
    /// - [`SparklineType::Column`]: A histogram style sparkline.
    ///
    ///   <img src="https://rustxlsxwriter.github.io/images/sparkline_type_column.png">
    ///
    /// - [`SparklineType::WinLose`]: A positive/negative style sparkline. It
    ///   looks similar to a histogram but all the bars are the same height,
    ///
    ///   <img src="https://rustxlsxwriter.github.io/images/sparkline_type_winlose.png">
    ///
    /// # Parameters
    ///
    /// - `sparkline_type`: A [`SparklineType`] value.
    ///
    pub fn set_type(mut self, sparkline_type: SparklineType) -> Sparkline {
        self.sparkline_type = sparkline_type;
        self
    }

    /// Display the highest point(s) in a sparkline with a marker.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_high_point.png">
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn show_high_point(mut self, enable: bool) -> Sparkline {
        self.show_high_point = enable;
        self
    }

    /// Display the lowest point(s) in a sparkline with a marker.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_low_point.png">
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn show_low_point(mut self, enable: bool) -> Sparkline {
        self.show_low_point = enable;
        self
    }

    /// Display the first point in a sparkline with a marker.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_first_point.png">
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn show_first_point(mut self, enable: bool) -> Sparkline {
        self.show_first_point = enable;
        self
    }

    /// Display the last point in a sparkline with a marker.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_last_point.png">
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn show_last_point(mut self, enable: bool) -> Sparkline {
        self.show_last_point = enable;
        self
    }

    /// Display the negative points in a sparkline with markers.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_negative_points.png">
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn show_negative_points(mut self, enable: bool) -> Sparkline {
        self.show_negative_points = enable;
        self
    }

    /// Display markers for all points in the sparkline.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_markers.png">
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn show_markers(mut self, enable: bool) -> Sparkline {
        self.show_markers = enable;
        self
    }

    /// Display the horizontal axis for a sparkline.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_axis.png">
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn show_axis(mut self, enable: bool) -> Sparkline {
        self.show_axis = enable;
        self
    }

    /// Display data from hidden rows or columns in a sparkline.
    ///

    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn show_hidden_data(mut self, enable: bool) -> Sparkline {
        self.show_hidden_data = enable;
        self
    }

    /// Set the option for displaying empty cells in a sparkline.
    ///
    /// The options are:
    ///
    /// - [`ChartEmptyCells::Gaps`]: Show empty cells in the chart as gaps. The
    ///   default.
    /// - [`ChartEmptyCells::Zero`]: Show empty cells in the chart as zeroes.
    /// - [`ChartEmptyCells::Connected`]: Show empty cells in the chart
    ///   connected by a line to the previous point.
    ///
    /// # Parameters
    ///
    /// `option` - A [`ChartEmptyCells`] enum value.
    ///
    pub fn show_empty_cells_as(mut self, option: ChartEmptyCells) -> Sparkline {
        self.show_empty_cells_as = option;

        self
    }

    /// Display the sparkline in right to left, reversed order.
    ///
    /// Change the default direction of the sparkline so that it is plotted from
    /// right to left instead of the default left to right.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn set_right_to_left(mut self, enable: bool) -> Sparkline {
        self.show_right_to_left = enable;
        self
    }

    /// Set the color of a sparkline.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a [`Color`] enum value or a
    ///   type that can convert [`Into`] a [`Color`] such as a html string.
    ///
    /// # Examples
    ///
    /// The following example demonstrates adding a sparkline to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_sparkline_set_sparkline_color.rs
    /// #
    /// # use rust_xlsxwriter::{Sparkline, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some sample data to plot.
    /// #     worksheet.write_row(0, 0, [-2, 2, 3, -1, 0])?;
    /// #
    ///     // Create a default line sparkline and set its color.
    ///     let sparkline = Sparkline::new()
    ///         .set_range(("Sheet1", 0, 0, 0, 4))
    ///         .set_sparkline_color("#CF6348");
    ///
    ///     // Add it to the worksheet.
    ///     worksheet.add_sparkline(0, 5, &sparkline)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("sparkline.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_set_sparkline_color.png">
    ///
    ///
    pub fn set_sparkline_color(mut self, color: impl Into<Color>) -> Sparkline {
        let color = color.into();
        if color.is_valid() {
            self.series_color = color;
        }
        self
    }

    /// Turn on and set the color the sparkline highest point marker.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a [`Color`] enum value or a
    ///   type that can convert [`Into`] a [`Color`] such as a html string.
    ///
    pub fn set_high_point_color(mut self, color: impl Into<Color>) -> Sparkline {
        let color = color.into();
        if color.is_valid() {
            self.high_point_color = color;
            self.show_high_point = true;
        }
        self
    }

    /// Turn on and set the color the sparkline lowest point marker.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a [`Color`] enum value or a
    ///   type that can convert [`Into`] a [`Color`] such as a html string.
    ///
    pub fn set_low_point_color(mut self, color: impl Into<Color>) -> Sparkline {
        let color = color.into();
        if color.is_valid() {
            self.low_point_color = color;
            self.show_low_point = true;
        }
        self
    }

    /// Turn on and set the color the sparkline first point marker.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a [`Color`] enum value or a
    ///   type that can convert [`Into`] a [`Color`] such as a html string.
    ///
    pub fn set_first_point_color(mut self, color: impl Into<Color>) -> Sparkline {
        let color = color.into();
        if color.is_valid() {
            self.first_point_color = color;
            self.show_first_point = true;
        }
        self
    }

    /// Turn on and set the color the sparkline last point marker.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a [`Color`] enum value or a
    ///   type that can convert [`Into`] a [`Color`] such as a html string.
    ///
    pub fn set_last_point_color(mut self, color: impl Into<Color>) -> Sparkline {
        let color = color.into();
        if color.is_valid() {
            self.last_point_color = color;
            self.show_last_point = true;
        }
        self
    }

    /// Turn on and set the color the sparkline negative point markers.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a [`Color`] enum value or a
    ///   type that can convert [`Into`] a [`Color`] such as a html string.
    ///
    pub fn set_negative_points_color(mut self, color: impl Into<Color>) -> Sparkline {
        let color = color.into();
        if color.is_valid() {
            self.negative_points_color = color;
            self.show_negative_points = true;
        }
        self
    }

    /// Turn on and set the color the sparkline point markers.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a [`Color`] enum value or a
    ///   type that can convert [`Into`] a [`Color`] such as a html string.
    ///
    pub fn set_markers_color(mut self, color: impl Into<Color>) -> Sparkline {
        let color = color.into();
        if color.is_valid() {
            self.markers_color = color;
            self.show_markers = true;
        }
        self
    }

    /// Set the weight/width of the sparkline line.
    ///
    /// # Parameters
    ///
    /// - `weight`: The weight/width of the sparkline line. The width can be an
    ///   number type that convert [`Into`] [`f64`]. The default is 0.75.
    ///
    pub fn set_line_weight<T>(mut self, weight: T) -> Sparkline
    where
        T: Into<f64>,
    {
        self.line_weight = Some(weight.into());
        self
    }

    /// Set the maximum vertical value for a sparkline.
    ///
    /// Set the maximum bound to be displayed for the vertical axis of a
    /// sparkline.
    ///
    /// # Parameters
    ///
    /// `max` - The maximum bound for the axes.
    ///
    pub fn set_custom_max<T>(mut self, max: T) -> Sparkline
    where
        T: Into<f64>,
    {
        self.custom_max = Some(max.into());
        self.group_max = false;
        self
    }

    /// Set the minimum vertical value for a sparkline.
    ///
    /// Set the minimum bound to be displayed for the vertical axis of a
    /// sparkline.
    ///
    /// # Parameters
    ///
    /// `min` - The minimum bound for the axes.
    ///
    pub fn set_custom_min<T>(mut self, min: T) -> Sparkline
    where
        T: Into<f64>,
    {
        self.custom_min = Some(min.into());
        self.group_min = false;
        self
    }

    /// Set the maximum vertical value for a group of sparklines.
    ///
    /// Set the maximum vertical value for a group of sparklines based on the
    /// maximum value of the group.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn set_group_max(mut self, enable: bool) -> Sparkline {
        self.group_max = enable;

        if enable {
            self.custom_max = None;
        }

        self
    }

    /// Set the minimum vertical value for a group of sparklines.
    ///
    /// Set the minimum vertical value for a group of sparklines based on the
    /// minimum value of the group.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn set_group_min(mut self, enable: bool) -> Sparkline {
        self.group_min = enable;

        if enable {
            self.custom_min = None;
        }

        self
    }

    /// Set the an optional date axis for the sparkline data.
    ///
    /// In general Excel graphs sparklines at equally spaced X intervals.
    /// However, it is also possible to specify an optional range of dates that
    /// can be used as the X values `set_date_range()`.
    ///
    /// # Parameters
    ///
    /// - `range`: A 1D range that contains the dates used to plot the
    ///   sparkline. This can specified in different ways, see
    ///   [`IntoChartRange`] for details.
    ///
    pub fn set_date_range<T>(mut self, range: T) -> Sparkline
    where
        T: IntoChartRange,
    {
        self.date_range = range.new_chart_range();
        self
    }

    /// Change the data range order for 2D data ranges in grouped sparklines.
    ///
    /// When creating grouped sparklines via the
    /// [`Worksheet::add_sparkline_group()`](crate::Worksheet::add_sparkline_group())
    /// method the data range that the sparkline is applied to is in row major
    /// order, i.e., row by row. If required you can change this to column major
    /// order using this method.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn set_column_order(mut self, enable: bool) -> Sparkline {
        self.data_row_order = !enable;
        self
    }

    /// Set the sparkline style type.
    ///
    /// The `set_style()` method is used to set the style of the sparkline to
    /// one of 36 built-in styles. The default style is 1. The image below shows
    /// the 36 default styles. The index is counted from the top left and then
    /// in column-row order.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_styles.png">
    ///
    /// # Parameters
    ///
    /// - `style`: A integer value in the range 1-36.
    ///
    #[allow(clippy::too_many_lines)]
    #[allow(clippy::unreadable_literal)]
    pub fn set_style(mut self, style: u8) -> Sparkline {
        match style {
            1 => {
                self.low_point_color = Color::Theme(4, 0);
                self.high_point_color = Color::Theme(4, 0);
                self.last_point_color = Color::Theme(4, 3);
                self.first_point_color = Color::Theme(4, 3);
                self.markers_color = Color::Theme(4, 5);
                self.negative_points_color = Color::Theme(5, 0);
                self.series_color = Color::Theme(4, 5);
            }
            2 => {
                self.low_point_color = Color::Theme(5, 0);
                self.high_point_color = Color::Theme(5, 0);
                self.last_point_color = Color::Theme(5, 3);
                self.first_point_color = Color::Theme(5, 3);
                self.markers_color = Color::Theme(5, 5);
                self.negative_points_color = Color::Theme(6, 0);
                self.series_color = Color::Theme(5, 5);
            }
            3 => {
                self.low_point_color = Color::Theme(6, 0);
                self.high_point_color = Color::Theme(6, 0);
                self.last_point_color = Color::Theme(6, 3);
                self.first_point_color = Color::Theme(6, 3);
                self.markers_color = Color::Theme(6, 5);
                self.negative_points_color = Color::Theme(7, 0);
                self.series_color = Color::Theme(6, 5);
            }
            4 => {
                self.low_point_color = Color::Theme(7, 0);
                self.high_point_color = Color::Theme(7, 0);
                self.last_point_color = Color::Theme(7, 3);
                self.first_point_color = Color::Theme(7, 3);
                self.markers_color = Color::Theme(7, 5);
                self.negative_points_color = Color::Theme(8, 0);
                self.series_color = Color::Theme(7, 5);
            }
            5 => {
                self.low_point_color = Color::Theme(8, 0);
                self.high_point_color = Color::Theme(8, 0);
                self.last_point_color = Color::Theme(8, 3);
                self.first_point_color = Color::Theme(8, 3);
                self.markers_color = Color::Theme(8, 5);
                self.negative_points_color = Color::Theme(9, 0);
                self.series_color = Color::Theme(8, 5);
            }
            6 => {
                self.low_point_color = Color::Theme(9, 0);
                self.high_point_color = Color::Theme(9, 0);
                self.last_point_color = Color::Theme(9, 3);
                self.first_point_color = Color::Theme(9, 3);
                self.markers_color = Color::Theme(9, 5);
                self.negative_points_color = Color::Theme(4, 0);
                self.series_color = Color::Theme(9, 5);
            }
            7 => {
                self.low_point_color = Color::Theme(5, 4);
                self.high_point_color = Color::Theme(5, 4);
                self.last_point_color = Color::Theme(5, 4);
                self.first_point_color = Color::Theme(5, 4);
                self.markers_color = Color::Theme(5, 4);
                self.negative_points_color = Color::Theme(5, 0);
                self.series_color = Color::Theme(4, 4);
            }
            8 => {
                self.low_point_color = Color::Theme(6, 4);
                self.high_point_color = Color::Theme(6, 4);
                self.last_point_color = Color::Theme(6, 4);
                self.first_point_color = Color::Theme(6, 4);
                self.markers_color = Color::Theme(6, 4);
                self.negative_points_color = Color::Theme(6, 0);
                self.series_color = Color::Theme(5, 4);
            }
            9 => {
                self.low_point_color = Color::Theme(7, 4);
                self.high_point_color = Color::Theme(7, 4);
                self.last_point_color = Color::Theme(7, 4);
                self.first_point_color = Color::Theme(7, 4);
                self.markers_color = Color::Theme(7, 4);
                self.negative_points_color = Color::Theme(7, 0);
                self.series_color = Color::Theme(6, 4);
            }
            10 => {
                self.low_point_color = Color::Theme(8, 4);
                self.high_point_color = Color::Theme(8, 4);
                self.last_point_color = Color::Theme(8, 4);
                self.first_point_color = Color::Theme(8, 4);
                self.markers_color = Color::Theme(8, 4);
                self.negative_points_color = Color::Theme(8, 0);
                self.series_color = Color::Theme(7, 4);
            }
            11 => {
                self.low_point_color = Color::Theme(9, 4);
                self.high_point_color = Color::Theme(9, 4);
                self.last_point_color = Color::Theme(9, 4);
                self.first_point_color = Color::Theme(9, 4);
                self.markers_color = Color::Theme(9, 4);
                self.negative_points_color = Color::Theme(9, 0);
                self.series_color = Color::Theme(8, 4);
            }
            12 => {
                self.low_point_color = Color::Theme(4, 4);
                self.high_point_color = Color::Theme(4, 4);
                self.last_point_color = Color::Theme(4, 4);
                self.first_point_color = Color::Theme(4, 4);
                self.markers_color = Color::Theme(4, 4);
                self.negative_points_color = Color::Theme(4, 0);
                self.series_color = Color::Theme(9, 4);
            }
            13 => {
                self.low_point_color = Color::Theme(4, 4);
                self.high_point_color = Color::Theme(4, 4);
                self.last_point_color = Color::Theme(4, 4);
                self.first_point_color = Color::Theme(4, 4);
                self.markers_color = Color::Theme(4, 4);
                self.negative_points_color = Color::Theme(5, 0);
                self.series_color = Color::Theme(4, 0);
            }
            14 => {
                self.low_point_color = Color::Theme(5, 4);
                self.high_point_color = Color::Theme(5, 4);
                self.last_point_color = Color::Theme(5, 4);
                self.first_point_color = Color::Theme(5, 4);
                self.markers_color = Color::Theme(5, 4);
                self.negative_points_color = Color::Theme(6, 0);
                self.series_color = Color::Theme(5, 0);
            }
            15 => {
                self.low_point_color = Color::Theme(6, 4);
                self.high_point_color = Color::Theme(6, 4);
                self.last_point_color = Color::Theme(6, 4);
                self.first_point_color = Color::Theme(6, 4);
                self.markers_color = Color::Theme(6, 4);
                self.negative_points_color = Color::Theme(7, 0);
                self.series_color = Color::Theme(6, 0);
            }
            16 => {
                self.low_point_color = Color::Theme(7, 4);
                self.high_point_color = Color::Theme(7, 4);
                self.last_point_color = Color::Theme(7, 4);
                self.first_point_color = Color::Theme(7, 4);
                self.markers_color = Color::Theme(7, 4);
                self.negative_points_color = Color::Theme(8, 0);
                self.series_color = Color::Theme(7, 0);
            }
            17 => {
                self.low_point_color = Color::Theme(8, 4);
                self.high_point_color = Color::Theme(8, 4);
                self.last_point_color = Color::Theme(8, 4);
                self.first_point_color = Color::Theme(8, 4);
                self.markers_color = Color::Theme(8, 4);
                self.negative_points_color = Color::Theme(9, 0);
                self.series_color = Color::Theme(8, 0);
            }
            18 => {
                self.low_point_color = Color::Theme(9, 4);
                self.high_point_color = Color::Theme(9, 4);
                self.last_point_color = Color::Theme(9, 4);
                self.first_point_color = Color::Theme(9, 4);
                self.markers_color = Color::Theme(9, 4);
                self.negative_points_color = Color::Theme(4, 0);
                self.series_color = Color::Theme(9, 0);
            }
            19 => {
                self.low_point_color = Color::Theme(4, 5);
                self.high_point_color = Color::Theme(4, 5);
                self.last_point_color = Color::Theme(4, 4);
                self.first_point_color = Color::Theme(4, 4);
                self.markers_color = Color::Theme(4, 1);
                self.negative_points_color = Color::Theme(0, 5);
                self.series_color = Color::Theme(4, 3);
            }
            20 => {
                self.low_point_color = Color::Theme(5, 5);
                self.high_point_color = Color::Theme(5, 5);
                self.last_point_color = Color::Theme(5, 4);
                self.first_point_color = Color::Theme(5, 4);
                self.markers_color = Color::Theme(5, 1);
                self.negative_points_color = Color::Theme(0, 5);
                self.series_color = Color::Theme(5, 3);
            }
            21 => {
                self.low_point_color = Color::Theme(6, 5);
                self.high_point_color = Color::Theme(6, 5);
                self.last_point_color = Color::Theme(6, 4);
                self.first_point_color = Color::Theme(6, 4);
                self.markers_color = Color::Theme(6, 1);
                self.negative_points_color = Color::Theme(0, 5);
                self.series_color = Color::Theme(6, 3);
            }
            22 => {
                self.low_point_color = Color::Theme(7, 5);
                self.high_point_color = Color::Theme(7, 5);
                self.last_point_color = Color::Theme(7, 4);
                self.first_point_color = Color::Theme(7, 4);
                self.markers_color = Color::Theme(7, 1);
                self.negative_points_color = Color::Theme(0, 5);
                self.series_color = Color::Theme(7, 3);
            }
            23 => {
                self.low_point_color = Color::Theme(8, 5);
                self.high_point_color = Color::Theme(8, 5);
                self.last_point_color = Color::Theme(8, 4);
                self.first_point_color = Color::Theme(8, 4);
                self.markers_color = Color::Theme(8, 1);
                self.negative_points_color = Color::Theme(0, 5);
                self.series_color = Color::Theme(8, 3);
            }
            24 => {
                self.low_point_color = Color::Theme(9, 5);
                self.high_point_color = Color::Theme(9, 5);
                self.last_point_color = Color::Theme(9, 4);
                self.first_point_color = Color::Theme(9, 4);
                self.markers_color = Color::Theme(9, 1);
                self.negative_points_color = Color::Theme(0, 5);
                self.series_color = Color::Theme(9, 3);
            }
            25 => {
                self.low_point_color = Color::Theme(1, 3);
                self.high_point_color = Color::Theme(1, 3);
                self.last_point_color = Color::Theme(1, 3);
                self.first_point_color = Color::Theme(1, 3);
                self.markers_color = Color::Theme(1, 3);
                self.negative_points_color = Color::Theme(1, 3);
                self.series_color = Color::Theme(1, 1);
            }
            26 => {
                self.low_point_color = Color::Theme(0, 3);
                self.high_point_color = Color::Theme(0, 3);
                self.last_point_color = Color::Theme(0, 3);
                self.first_point_color = Color::Theme(0, 3);
                self.markers_color = Color::Theme(0, 3);
                self.negative_points_color = Color::Theme(0, 3);
                self.series_color = Color::Theme(1, 2);
            }
            27 => {
                self.low_point_color = Color::RGB(0xD00000);
                self.high_point_color = Color::RGB(0xD00000);
                self.last_point_color = Color::RGB(0xD00000);
                self.first_point_color = Color::RGB(0xD00000);
                self.markers_color = Color::RGB(0xD00000);
                self.negative_points_color = Color::RGB(0xD00000);
                self.series_color = Color::RGB(0x323232);
            }
            28 => {
                self.low_point_color = Color::RGB(0x00070C0);
                self.high_point_color = Color::RGB(0x00070C0);
                self.last_point_color = Color::RGB(0x00070C0);
                self.first_point_color = Color::RGB(0x00070C0);
                self.markers_color = Color::RGB(0x00070C0);
                self.negative_points_color = Color::RGB(0x00070C0);
                self.series_color = Color::RGB(0x000000);
            }
            29 => {
                self.low_point_color = Color::RGB(0xD00000);
                self.high_point_color = Color::RGB(0xD00000);
                self.last_point_color = Color::RGB(0xD00000);
                self.first_point_color = Color::RGB(0xD00000);
                self.markers_color = Color::RGB(0xD00000);
                self.negative_points_color = Color::RGB(0xD00000);
                self.series_color = Color::RGB(0x376092);
            }
            30 => {
                self.low_point_color = Color::RGB(0x000000);
                self.high_point_color = Color::RGB(0x000000);
                self.last_point_color = Color::RGB(0x000000);
                self.first_point_color = Color::RGB(0x000000);
                self.markers_color = Color::RGB(0x000000);
                self.negative_points_color = Color::RGB(0x000000);
                self.series_color = Color::RGB(0x00070C0);
            }
            31 => {
                self.low_point_color = Color::RGB(0xFF5055);
                self.high_point_color = Color::RGB(0x56BE79);
                self.last_point_color = Color::RGB(0x359CEB);
                self.first_point_color = Color::RGB(0x5687C2);
                self.markers_color = Color::RGB(0xD70077);
                self.negative_points_color = Color::RGB(0xFFB620);
                self.series_color = Color::RGB(0x5F5F5F);
            }
            32 => {
                self.low_point_color = Color::RGB(0xFF5055);
                self.high_point_color = Color::RGB(0x56BE79);
                self.last_point_color = Color::RGB(0x359CEB);
                self.first_point_color = Color::RGB(0x777777);
                self.markers_color = Color::RGB(0xD70077);
                self.negative_points_color = Color::RGB(0xFFB620);
                self.series_color = Color::RGB(0x5687C2);
            }
            33 => {
                self.low_point_color = Color::RGB(0xFF5367);
                self.high_point_color = Color::RGB(0x60D276);
                self.last_point_color = Color::RGB(0xFFEB9C);
                self.first_point_color = Color::RGB(0xFFDC47);
                self.markers_color = Color::RGB(0x8CADD6);
                self.negative_points_color = Color::RGB(0xFFC7CE);
                self.series_color = Color::RGB(0xC6EFCE);
            }
            34 => {
                self.low_point_color = Color::RGB(0xFF0000);
                self.high_point_color = Color::RGB(0x00B050);
                self.last_point_color = Color::RGB(0xFFC000);
                self.first_point_color = Color::RGB(0xFFC000);
                self.markers_color = Color::RGB(0x0070C0);
                self.negative_points_color = Color::RGB(0xFF0000);
                self.series_color = Color::RGB(0x0B050);
            }
            35 => {
                self.low_point_color = Color::Theme(7, 0);
                self.high_point_color = Color::Theme(6, 0);
                self.last_point_color = Color::Theme(5, 0);
                self.first_point_color = Color::Theme(4, 0);
                self.markers_color = Color::Theme(8, 0);
                self.negative_points_color = Color::Theme(9, 0);
                self.series_color = Color::Theme(3, 0);
            }
            36 => {
                self.low_point_color = Color::Theme(7, 0);
                self.high_point_color = Color::Theme(6, 0);
                self.last_point_color = Color::Theme(5, 0);
                self.first_point_color = Color::Theme(4, 0);
                self.markers_color = Color::Theme(8, 0);
                self.negative_points_color = Color::Theme(9, 0);
                self.series_color = Color::Theme(1, 0);
            }
            _ => eprintln!("Sparkline style '{style}' outside the Excel range 1-36."),
        };

        self
    }

    // Add a single sparkline to a cell.
    pub(crate) fn add_cell_range(&mut self, row: RowNum, col: ColNum) {
        let cell = utility::row_col_to_cell(row, col);
        let range = self.data_range.formula();

        self.ranges.push((cell, range));
    }

    // Add a group sparkline to a range.
    pub(crate) fn add_group_range(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
    ) {
        let cell_row_order = last_col - first_col == 0;
        self.data_range.set_baseline(self.data_row_order);

        if cell_row_order {
            for row in first_row..=last_row {
                let cell = utility::row_col_to_cell(row, first_col);
                let range = self.data_range.formula();
                self.ranges.push((cell, range));
                self.data_range.increment(self.data_row_order);
            }
        } else {
            for col in first_col..=last_col {
                let cell = utility::row_col_to_cell(first_row, col);
                let range = self.data_range.formula();
                self.ranges.push((cell, range));
                self.data_range.increment(self.data_row_order);
            }
        }
    }
}

// -----------------------------------------------------------------------
// Helper enums/structs/traits
// -----------------------------------------------------------------------

/// The `SparklineType` enum defines [`Sparkline`] types.
///
/// This is used with the [`Sparkline::set_type()`](Sparkline::set_type())
/// method.
///
#[derive(Clone, Copy, Eq, PartialEq)]
pub enum SparklineType {
    /// A line style sparkline. This is the default.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_type_line.png">
    ///
    Line,

    /// A histogram style sparkline.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_type_column.png">
    ///
    Column,

    /// A positive/negative style sparkline. It looks similar to a histogram but
    /// all the bars are the same height,
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_type_winlose.png">
    ///
    WinLose,
}
