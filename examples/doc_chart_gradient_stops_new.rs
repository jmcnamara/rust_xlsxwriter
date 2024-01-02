// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of creating gradient stops for a gradient fill for a chart element.

use rust_xlsxwriter::ChartGradientStop;

#[allow(unused_variables)]
fn main() {
    let gradient_stops = [
        ChartGradientStop::new("#156B13", 0),
        ChartGradientStop::new("#9CB86E", 50),
        ChartGradientStop::new("#DDEBCF", 100),
    ];
}
