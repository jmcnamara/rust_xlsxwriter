// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates adding a note to a worksheet cell. This
//! example also changes the note dimensions.

use rust_xlsxwriter::{Note, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    // Create a new note.
    let note = Note::new("Some text for the note")
        .set_width(200)
        .set_height(200);

    // Add the note to a worksheet cell.
    worksheet.insert_note(2, 0, &note)?;

    // Save the file to disk.
    workbook.save("notes.xlsx")?;

    Ok(())
}
