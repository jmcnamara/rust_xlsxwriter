// Sparkline - A module to represent an Excel sparkline.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use crate::{ChartRange, Color, IntoChartRange};

mod tests;

/// TODO
#[derive(Clone)]
pub struct Sparkline {
    pub(crate) has_gap: bool,
    pub(crate) series_color: Color,
    pub(crate) negative_color: Color,
    pub(crate) axis_color: Color,
    pub(crate) marker_color: Color,
    pub(crate) first_color: Color,
    pub(crate) last_color: Color,
    pub(crate) high_color: Color,
    pub(crate) low_color: Color,
    pub(crate) range: ChartRange,
    pub(crate) cell: String,
}

#[allow(clippy::new_without_default)]
impl Sparkline {
    /// Create a new Sparkline struct.
    pub fn new() -> Sparkline {
        Sparkline {
            has_gap: true,
            series_color: Color::Theme(4, 5),
            negative_color: Color::Theme(5, 0),
            axis_color: Color::Black,
            marker_color: Color::Theme(4, 5),
            first_color: Color::Theme(4, 3),
            last_color: Color::Theme(4, 3),
            high_color: Color::Theme(4, 0),
            low_color: Color::Theme(4, 0),
            range: ChartRange::default(),
            cell: String::new(),
        }
    }

    /// TODO
    pub fn set_range<T>(mut self, range: T) -> Sparkline
    where
        T: IntoChartRange,
    {
        self.range = range.new_chart_range();

        self
    }
}
