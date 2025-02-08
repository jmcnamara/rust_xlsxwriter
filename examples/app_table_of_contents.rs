// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! This is an example of creating a "Table of Contents" worksheet with links to
//! other worksheets in the workbook.

use rust_xlsxwriter::{utility::quote_sheet_name, Format, Url, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create a table of contents worksheet at the start. If the worksheet names
    // are known in advance you can do add them here. For the sake of this
    // example we will assume that they aren't known and/or are created
    // dynamically.
    let _ = workbook.add_worksheet().set_name("Overview")?;

    // Add some worksheets.
    let _ = workbook.add_worksheet().set_name("Pricing")?;
    let _ = workbook.add_worksheet().set_name("Sales")?;
    let _ = workbook.add_worksheet().set_name("Revenue")?;
    let _ = workbook.add_worksheet().set_name("Analytics")?;

    // If the sheet names aren't known in advance we can find them as follows:
    let mut worksheet_names = workbook
        .worksheets()
        .iter()
        .map(|worksheet| worksheet.name())
        .collect::<Vec<_>>();

    // Remove the "Overview" worksheet name.
    worksheet_names.remove(0);

    // Get the "Overview" worksheet to add the table of contents.
    let worksheet = workbook.worksheet_from_name("Overview")?;

    // Write a header.
    let header = Format::new().set_bold().set_background_color("C6EFCE");
    worksheet.write_string_with_format(0, 0, "Table of Contents", &header)?;

    // Write the worksheet names with links.
    for (i, name) in worksheet_names.iter().enumerate() {
        let sheet_name = quote_sheet_name(name);
        let link = format!("internal:{sheet_name}!A1");
        let url = Url::new(link).set_text(name);

        worksheet.write_url(i as u32 + 1, 0, &url)?;
    }

    // Autofit the data for clarity.
    worksheet.autofit();

    // Save the file to disk.
    workbook.save("table_of_contents.xlsx")?;

    Ok(())
}
