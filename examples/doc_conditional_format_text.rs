// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! Example of adding a text type conditional formatting to a worksheet.

use rust_xlsxwriter::{
    ConditionalFormatText, ConditionalFormatTextCriteria, Format, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some sample data.
    let data = [
        "apocrustic",
        "burstwort",
        "cloudburst",
        "crustification",
        "distrustfulness",
        "laurustine",
        "outburst",
        "rusticism",
        "thunderburst",
        "trustee",
        "trustworthiness",
        "unburstableness",
        "unfrustratable",
    ];
    worksheet.write_column(0, 0, data)?;
    worksheet.write_column(0, 2, data)?;

    // Set the column widths for clarity.
    worksheet.set_column_width(0, 20)?;
    worksheet.set_column_width(2, 20)?;

    // Add a format. Green fill with dark green text.
    let format1 = Format::new()
        .set_font_color("006100")
        .set_background_color("C6EFCE");

    // Add a format. Light red fill with dark red text.
    let format2 = Format::new()
        .set_font_color("9C0006")
        .set_background_color("FFC7CE");

    // Write a text "containing" conditional format over a range.
    let conditional_format = ConditionalFormatText::new()
        .set_criteria(ConditionalFormatTextCriteria::Contains)
        .set_value("rust")
        .set_format(&format1);

    worksheet.add_conditional_format(0, 0, 12, 0, &conditional_format)?;

    // Write a text "not containing" conditional format over the same range.
    let conditional_format = ConditionalFormatText::new()
        .set_criteria(ConditionalFormatTextCriteria::DoesNotContain)
        .set_value("rust")
        .set_format(&format2);

    worksheet.add_conditional_format(0, 0, 12, 0, &conditional_format)?;

    // Write a text "begins with" conditional format over a range.
    let conditional_format = ConditionalFormatText::new()
        .set_criteria(ConditionalFormatTextCriteria::BeginsWith)
        .set_value("t")
        .set_format(&format1);

    worksheet.add_conditional_format(0, 2, 12, 2, &conditional_format)?;

    // Write a text "ends with" conditional format over the same range.
    let conditional_format = ConditionalFormatText::new()
        .set_criteria(ConditionalFormatTextCriteria::EndsWith)
        .set_value("t")
        .set_format(&format2);

    worksheet.add_conditional_format(0, 2, 12, 2, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
