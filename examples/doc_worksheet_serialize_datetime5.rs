// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! Example of a serializable struct with an Option Chrono Naive value with a
//! helper function.
//!
use chrono::NaiveDate;
use rust_xlsxwriter::utility::serialize_option_datetime_to_excel;
use serde::Serialize;

fn main() {
    #[derive(Serialize)]
    struct Student {
        full_name: String,

        #[serde(serialize_with = "serialize_option_datetime_to_excel")]
        birth_date: Option<NaiveDate>,

        id_number: u32,
    }
}
