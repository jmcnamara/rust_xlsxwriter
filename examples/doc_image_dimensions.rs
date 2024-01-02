// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! This example shows how to get some of the properties of an Image that will
//! be used in an Excel worksheet.

use rust_xlsxwriter::{Image, XlsxError};

fn main() -> Result<(), XlsxError> {
    let image = Image::new("examples/rust_logo.png")?;

    assert_eq!(106.0, image.width());
    assert_eq!(106.0, image.height());
    assert_eq!(96.0, image.width_dpi());
    assert_eq!(96.0, image.height_dpi());

    Ok(())
}
