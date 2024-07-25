// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates adding notes to a worksheet and setting
//! the worksheet property to make them all visible.

use rust_xlsxwriter::{Note, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    // Create a new note.
    let note = Note::new("Some text for the note");

    // Add the note to some worksheet cells.
    worksheet.insert_note(1, 0, &note)?;
    worksheet.insert_note(3, 3, &note)?;
    worksheet.insert_note(6, 0, &note)?;
    worksheet.insert_note(8, 3, &note)?;

    // Display all the notes in the worksheet.
    worksheet.show_all_notes(true);

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
