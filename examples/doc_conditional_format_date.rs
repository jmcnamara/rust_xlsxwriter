// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! Example of adding a Dates Occurring type conditional formatting to a
//! worksheet. Note, the rules in this example such as "Last month", "This
//! month" and "Next month" are applied to the sample dates which by default are
//! for November 2023. Changes the dates to some range closer to the time you
//! run the example.

use rust_xlsxwriter::{
    ConditionalFormatDate, ConditionalFormatDateCriteria, ExcelDateTime, Format, Workbook,
    XlsxError,
};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Create a date format.
    let date_format = Format::new().set_num_format("yyyy-mm-dd");

    // Add some sample data.
    let dates = [
        "2023-10-01",
        "2023-11-05",
        "2023-11-06",
        "2023-10-04",
        "2023-11-11",
        "2023-11-02",
        "2023-11-04",
        "2023-11-12",
        "2023-12-01",
        "2023-12-13",
        "2023-11-13",
    ];

    // Map the string dates to ExcelDateTime objects, while capturing any
    // potential conversion errors.
    let dates: Result<Vec<ExcelDateTime>, XlsxError> = dates
        .into_iter()
        .map(ExcelDateTime::parse_from_str)
        .collect();
    let dates = dates?;

    worksheet.write_column_with_format(0, 0, dates, &date_format)?;

    // Set the column widths for clarity.
    worksheet.set_column_width(0, 20)?;

    // Add a format. Light red fill with dark red text.
    let format1 = Format::new()
        .set_font_color("9C0006")
        .set_background_color("FFC7CE");

    // Add a format. Green fill with dark green text.
    let format2 = Format::new()
        .set_font_color("006100")
        .set_background_color("C6EFCE");

    // Add a format. Light yellow fill with dark yellow text.
    let format3 = Format::new()
        .set_font_color("9C6500")
        .set_background_color("FFEB9C");

    // Write conditional format over the same range.
    let conditional_format = ConditionalFormatDate::new()
        .set_criteria(ConditionalFormatDateCriteria::LastMonth)
        .set_format(format1);

    worksheet.add_conditional_format(0, 0, 10, 0, &conditional_format)?;

    let conditional_format = ConditionalFormatDate::new()
        .set_criteria(ConditionalFormatDateCriteria::ThisMonth)
        .set_format(format2);

    worksheet.add_conditional_format(0, 0, 10, 0, &conditional_format)?;

    let conditional_format = ConditionalFormatDate::new()
        .set_criteria(ConditionalFormatDateCriteria::NextMonth)
        .set_format(format3);

    worksheet.add_conditional_format(0, 0, 10, 0, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
