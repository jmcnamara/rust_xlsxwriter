// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing cell Notes to a worksheet using the rust_xlsxwriter
//! library.

use rust_xlsxwriter::{Note, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Widen the first column for clarity.
    worksheet.set_column_width(0, 16)?;

    // Write some data.
    let party_items = [
        "Invitations",
        "Doors",
        "Flowers",
        "Champagne",
        "Menu",
        "Peter",
    ];
    worksheet.write_column(0, 0, party_items)?;

    // Create a new worksheet Note.
    let note = Note::new("I will get the flowers myself").set_author("Clarissa Dalloway");

    // Add the note to a cell.
    worksheet.insert_note(2, 0, &note)?;

    // Save the file to disk.
    workbook.save("notes.xlsx")?;

    Ok(())
}
