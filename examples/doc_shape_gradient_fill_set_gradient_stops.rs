// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! This example demonstrates adding a Textbox shape and setting some of the
//! gradient fill properties.

use rust_xlsxwriter::{Shape, ShapeGradientFill, ShapeGradientStop, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set the properties of the gradient stops.
    let gradient_stops = [
        ShapeGradientStop::new("#DDEBCF", 0),
        ShapeGradientStop::new("#9CB86E", 50),
        ShapeGradientStop::new("#156B13", 100),
    ];

    // Create a textbox shape with formatting.
    let textbox = Shape::textbox()
        .set_text("This is some text")
        .set_format(&ShapeGradientFill::new().set_gradient_stops(&gradient_stops));

    // Insert a textbox in a cell.
    worksheet.insert_shape(1, 1, &textbox)?;

    // Save the file to disk.
    workbook.save("shape.xlsx")?;

    Ok(())
}
