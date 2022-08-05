// error - error values for the rust_xlsxwriter library.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use std::error::Error;
use std::fmt;

#[derive(Debug)]
pub enum XlsxError {
    RowColRange,
}

impl Error for XlsxError {}

impl fmt::Display for XlsxError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            XlsxError::RowColRange => write!(
                f,
                "Row or column exceeds Excel's allowed range (1,048,576 x 16,384)"
            ),
        }
    }
}
