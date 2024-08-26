// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of passing shape formatting parameters via the
//! [`IntoShapeFormat`] trait.

use rust_xlsxwriter::{Shape, ShapeFormat, ShapeSolidFill, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a formatted shape via ShapeFormat and ShapeSolidFill.
    let textbox = Shape::textbox().set_text("This is some text").set_format(
        &ShapeFormat::new().set_solid_fill(&ShapeSolidFill::new().set_color("#8ED154")),
    );

    // Insert a textbox in a cell.
    worksheet.insert_shape(1, 1, &textbox)?;

    // Create a formatted shape via ShapeSolidFill directly.
    let textbox = Shape::textbox()
        .set_text("This is some text")
        .set_format(&ShapeSolidFill::new().set_color("#8ED154"));

    // Insert a textbox in a cell.
    worksheet.insert_shape(1, 5, &textbox)?;

    // Save the file to disk.
    workbook.save("shape.xlsx")?;

    Ok(())
}
