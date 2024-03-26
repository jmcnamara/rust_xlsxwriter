// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example shows Excels default icon settings expressed as
//! `rust_xlsxwriter` rules.
//!
use rust_xlsxwriter::{ConditionalFormatCustomIcon, ConditionalFormatType};

#[allow(unused_variables)]
fn main() {
    // Default rules for three symbol icon sets.
    let icons3 = [
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 33),
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 67),
    ];

    // Default rules for four symbol icon sets.
    let icons4 = [
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 25),
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 50),
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 75),
    ];

    // Default rules for five symbol icon sets.
    let icons5 = [
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 20),
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 40),
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 60),
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 80),
    ];
}
