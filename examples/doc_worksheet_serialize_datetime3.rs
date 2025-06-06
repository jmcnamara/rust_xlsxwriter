// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! Example of a serializable struct with a Chrono Naive value with a helper
//! function.

use chrono::NaiveDate;
use serde::Serialize;

use rust_xlsxwriter::utility::serialize_datetime_to_excel;

fn main() {
    #[derive(Serialize)]
    struct Student {
        full_name: String,

        #[serde(serialize_with = "serialize_datetime_to_excel")]
        birth_date: NaiveDate,

        id_number: u32,
    }
}
