// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! This example demonstrates adding a Textbox shape and setting some of the
//! line properties.

use rust_xlsxwriter::{Shape, ShapeFormat, ShapeLine, ShapeLineDashType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a textbox shape with formatting.
    let textbox = Shape::textbox().set_text("This is some text").set_format(
        &ShapeFormat::new().set_line(
            &ShapeLine::new()
                .set_color("#FF0000")
                .set_dash_type(ShapeLineDashType::DashDot),
        ),
    );

    // Insert a textbox in a cell.
    worksheet.insert_shape(1, 1, &textbox)?;

    // Save the file to disk.
    workbook.save("shape.xlsx")?;

    Ok(())
}
