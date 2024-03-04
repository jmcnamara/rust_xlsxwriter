// Sparkline - A module to represent an Excel sparkline.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use crate::{ChartEmptyCells, ChartRange, Color, IntoChartRange, IntoColor};

mod tests;

/// TODO
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
    pub(crate) range: ChartRange,
    pub(crate) date_range: ChartRange,
    pub(crate) cell: String,
    pub(crate) sparkline_type: SparklineType,
    pub(crate) show_high_point: bool,
    pub(crate) show_low_point: bool,
    pub(crate) show_first_point: bool,
    pub(crate) show_last_point: bool,
    pub(crate) show_negative_points: bool,
    pub(crate) show_markers: bool,
    pub(crate) show_axis: bool,
    pub(crate) show_hidden_data: bool,
    pub(crate) show_reversed: bool,
    pub(crate) show_empty_cells_as: ChartEmptyCells,
    pub(crate) line_weight: Option<f64>,
    pub(crate) custom_min: Option<f64>,
    pub(crate) custom_max: Option<f64>,
    pub(crate) group_max: bool,
    pub(crate) group_min: bool,
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
            range: ChartRange::default(),
            date_range: ChartRange::default(),
            cell: String::new(),
            sparkline_type: SparklineType::Line,
            show_high_point: false,
            show_low_point: false,
            show_first_point: false,
            show_last_point: false,
            show_negative_points: false,
            show_markers: false,
            show_axis: false,
            show_hidden_data: false,
            show_reversed: false,
            show_empty_cells_as: ChartEmptyCells::Gaps,
            line_weight: None,
            custom_min: None,
            custom_max: None,
            group_max: false,
            group_min: false,
        }
    }

    /// Set the range to which the sparkline applies.
    ///
    /// Excel graphs sparklines for using a user specified range as the Y values
    /// for the vertical axis of the plot. It uses evenly spaced X values for
    /// the horizontal axis.
    ///
    /// It is also possible to use a user specified range of dates for the X
    /// values using the
    /// [`Sparkline::set_date_range`](Sparkline::set_date_range) method.
    ///
    ///
    ///
    pub fn set_range<T>(mut self, range: T) -> Sparkline
    where
        T: IntoChartRange,
    {
        self.range = range.new_chart_range();
        self
    }

    /// Set the type of sparkline.
    ///
    /// TODO
    ///
    pub fn set_type(mut self, sparkline_type: SparklineType) -> Sparkline {
        self.sparkline_type = sparkline_type;
        self
    }

    /// Display the highest point in a sparkline with a marker.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn show_high_point(mut self, enable: bool) -> Sparkline {
        self.show_high_point = enable;
        self
    }

    /// Display the lowest point in a sparkline with a marker.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn show_low_point(mut self, enable: bool) -> Sparkline {
        self.show_low_point = enable;
        self
    }

    /// Display the first point in a sparkline with a marker.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn show_first_point(mut self, enable: bool) -> Sparkline {
        self.show_first_point = enable;
        self
    }

    /// Display the last point in a sparkline with a marker.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn show_last_point(mut self, enable: bool) -> Sparkline {
        self.show_last_point = enable;
        self
    }

    /// Display the negative points in a sparkline with markers.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn show_negative_points(mut self, enable: bool) -> Sparkline {
        self.show_negative_points = enable;
        self
    }

    /// Display markers for all points in the sparkline.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn show_markers(mut self, enable: bool) -> Sparkline {
        self.show_markers = enable;
        self
    }

    /// Display the horizontal axis for a sparkline.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn show_axis(mut self, enable: bool) -> Sparkline {
        self.show_axis = enable;
        self
    }

    /// Display data from hidden rows or columns in a sparkline.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn show_hidden_data(mut self, enable: bool) -> Sparkline {
        self.show_hidden_data = enable;
        self
    }

    /// Set the option for displaying empty cells in a sparkline.
    ///
    /// # Parameters
    ///
    /// `option` - A [`ChartEmptyCells`] enum value.
    ///
    pub fn show_empty_cells_as(mut self, option: ChartEmptyCells) -> Sparkline {
        self.show_empty_cells_as = option;

        self
    }

    /// Display the sparkline in reversed order.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn set_reverse(mut self, enable: bool) -> Sparkline {
        self.show_reversed = enable;
        self
    }

    /// Set the color of a sparkline.
    ///
    /// # Parameters
    ///
    /// * `color` - The color property defined by a [`Color`] enum value or a
    ///   type that implements the [`IntoColor`] trait such as a html string.
    ///
    pub fn set_sparkline_color<T>(mut self, color: T) -> Sparkline
    where
        T: IntoColor,
    {
        let color = color.new_color();
        if color.is_valid() {
            self.series_color = color;
        }
        self
    }

    /// Turn on and set the color the sparkline highest point marker.
    ///
    /// # Parameters
    ///
    /// * `color` - The color property defined by a [`Color`] enum value or a
    ///   type that implements the [`IntoColor`] trait such as a html string.
    ///
    pub fn set_high_point_color<T>(mut self, color: T) -> Sparkline
    where
        T: IntoColor,
    {
        let color = color.new_color();
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
    /// * `color` - The color property defined by a [`Color`] enum value or a
    ///   type that implements the [`IntoColor`] trait such as a html string.
    ///
    pub fn set_low_point_color<T>(mut self, color: T) -> Sparkline
    where
        T: IntoColor,
    {
        let color = color.new_color();
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
    /// * `color` - The color property defined by a [`Color`] enum value or a
    ///   type that implements the [`IntoColor`] trait such as a html string.
    ///
    pub fn set_first_point_color<T>(mut self, color: T) -> Sparkline
    where
        T: IntoColor,
    {
        let color = color.new_color();
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
    /// * `color` - The color property defined by a [`Color`] enum value or a
    ///   type that implements the [`IntoColor`] trait such as a html string.
    ///
    pub fn set_last_point_color<T>(mut self, color: T) -> Sparkline
    where
        T: IntoColor,
    {
        let color = color.new_color();
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
    /// * `color` - The color property defined by a [`Color`] enum value or a
    ///   type that implements the [`IntoColor`] trait such as a html string.
    ///
    pub fn set_negative_points_color<T>(mut self, color: T) -> Sparkline
    where
        T: IntoColor,
    {
        let color = color.new_color();
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
    /// * `color` - The color property defined by a [`Color`] enum value or a
    ///   type that implements the [`IntoColor`] trait such as a html string.
    ///
    pub fn set_markers_color<T>(mut self, color: T) -> Sparkline
    where
        T: IntoColor,
    {
        let color = color.new_color();
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
    /// * `weight` - The weight/width of the sparkline line. The width can be an
    /// number type that convert [`Into`] [`f64`]. The default is 0.75.
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
    /// maximum value for the group.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
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
    /// minimum value for the group.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    pub fn set_group_min(mut self, enable: bool) -> Sparkline {
        self.group_min = enable;

        if enable {
            self.custom_min = None;
        }

        self
    }

    /// Set the an option date axis for the sparkline data.
    ///
    /// In general Excel graphs sparklines at equally spaced X intervals. However,
    /// it is also possible to specify an optional range of dates that can be
    /// used as the X values `set_date_range()`.
    ///
    ///
    /// TODO
    ///
    pub fn set_date_range<T>(mut self, range: T) -> Sparkline
    where
        T: IntoChartRange,
    {
        self.date_range = range.new_chart_range();
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
    /// * `style` - A integer value in the range 1-36.
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
    Line,

    /// A histogram style sparkline.
    Column,

    /// A positive/negative style sparkline. It looks similar to a histogram but
    /// all the points are the same height,
    WinLose,
}
