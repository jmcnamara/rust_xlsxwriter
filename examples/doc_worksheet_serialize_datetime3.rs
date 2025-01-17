// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! Example of a serializable struct with a Chrono Naive value with a helper
//! function.

use chrono::NaiveDate;
use rust_xlsxwriter::utility::serialize_chrono_naive_to_excel;
use serde::Serialize;

fn main() {
    #[derive(Serialize)]
    struct Student {
        full_name: String,

        #[serde(serialize_with = "serialize_chrono_naive_to_excel")]
        birth_date: NaiveDate,

        id_number: u32,
    }
}
