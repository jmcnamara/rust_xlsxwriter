// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating a new chart range.
//!
use rust_xlsxwriter::ChartRange;

#[allow(unused_variables)]
fn main() {
    let range = ChartRange::new_from_string("Sheet1!$A$1:$A$5");
}
