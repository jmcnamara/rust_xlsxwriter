// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of how to add data validation and dropdown lists using the
//! rust_xlsxwriter library.
//!
//! Data validation is a feature of Excel which allows you to restrict the data
//! that a user enters in a cell and to display help and warning messages. It
//! also allows you to restrict input to values in a drop down list.

use rust_xlsxwriter::{
    DataValidation, DataValidationErrorStyle, DataValidationRule, ExcelDateTime, Format,
    FormatAlign, FormatBorder, Formula, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add a format for the header cells.
    let header_format = Format::new()
        .set_background_color("C6EFCE")
        .set_border(FormatBorder::Thin)
        .set_bold()
        .set_indent(1)
        .set_text_wrap()
        .set_align(FormatAlign::VerticalCenter);

    // Set up layout of the worksheet.
    worksheet.set_column_width(0, 68)?;
    worksheet.set_column_width(1, 15)?;
    worksheet.set_column_width(3, 15)?;
    worksheet.set_row_height(0, 36)?;

    // Write the header cells and some data that will be used in the examples.
    let heading1 = "Some examples of data validations";
    let heading2 = "Enter values in this column";
    let heading3 = "Sample Data";

    worksheet.write_with_format(0, 0, heading1, &header_format)?;
    worksheet.write_with_format(0, 1, heading2, &header_format)?;
    worksheet.write_with_format(0, 3, heading3, &header_format)?;

    worksheet.write(2, 3, "Integers")?;
    worksheet.write(2, 4, 1)?;
    worksheet.write(2, 5, 10)?;

    worksheet.write_row(3, 3, ["List data", "open", "high", "close"])?;

    worksheet.write(4, 3, "Formula")?;
    worksheet.write(4, 4, Formula::new("=AND(F5=50,G5=60)"))?;
    worksheet.write(4, 5, 50)?;
    worksheet.write(4, 6, 60)?;

    // -----------------------------------------------------------------------
    // Example 1. Limiting input to an integer in a fixed range.
    // -----------------------------------------------------------------------
    let text = "Enter an integer between 1 and 10";
    worksheet.write(2, 0, text)?;

    let data_validation =
        DataValidation::new().allow_whole_number(DataValidationRule::Between(1, 10));

    worksheet.add_data_validation(2, 1, 2, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 2. Limiting input to an integer outside a fixed range.
    // -----------------------------------------------------------------------
    let text = "Enter an integer that is not between 1 and 10 (using cell references)";
    worksheet.write(4, 0, text)?;

    let data_validation = DataValidation::new()
        .allow_whole_number_formula(DataValidationRule::NotBetween("=E3".into(), "=F3".into()));

    worksheet.add_data_validation(4, 1, 4, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 3. Limiting input to an integer greater than a fixed value.
    // -----------------------------------------------------------------------
    let text = "Enter an integer greater than 0";
    worksheet.write(6, 0, text)?;

    let data_validation =
        DataValidation::new().allow_whole_number(DataValidationRule::GreaterThan(0));

    worksheet.add_data_validation(6, 1, 6, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 4. Limiting input to an integer less than a fixed value.
    // -----------------------------------------------------------------------
    let text = "Enter an integer less than 10";
    worksheet.write(8, 0, text)?;

    let data_validation =
        DataValidation::new().allow_whole_number(DataValidationRule::LessThan(10));

    worksheet.add_data_validation(8, 1, 8, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 5. Limiting input to a decimal in a fixed range.
    // -----------------------------------------------------------------------
    let text = "Enter a decimal between 0.1 and 0.5";
    worksheet.write(10, 0, text)?;

    let data_validation =
        DataValidation::new().allow_decimal_number(DataValidationRule::Between(0.1, 0.5));

    worksheet.add_data_validation(10, 1, 10, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 6. Limiting input to a value in a dropdown list.
    // -----------------------------------------------------------------------
    let text = "Select a value from a drop down list";
    worksheet.write(12, 0, text)?;

    let data_validation = DataValidation::new().allow_list_strings(&["open", "high", "close"])?;

    worksheet.add_data_validation(12, 1, 12, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 7. Limiting input to a value in a dropdown list.
    // -----------------------------------------------------------------------
    let text = "Select a value from a drop down list (using a cell range)";
    worksheet.write(14, 0, text)?;

    let data_validation = DataValidation::new().allow_list_formula("=$E$4:$G$4".into());

    worksheet.add_data_validation(14, 1, 14, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 8. Limiting input to a date in a fixed range.
    // -----------------------------------------------------------------------
    let text = "Enter a date between 1/1/2025 and 12/12/2025";
    worksheet.write(16, 0, text)?;

    let data_validation = DataValidation::new().allow_date(DataValidationRule::Between(
        ExcelDateTime::parse_from_str("2025-01-01")?,
        ExcelDateTime::parse_from_str("2025-12-12")?,
    ));

    worksheet.add_data_validation(16, 1, 16, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 9. Limiting input to a time in a fixed range.
    // -----------------------------------------------------------------------
    let text = "Enter a time between 6:00 and 12:00";
    worksheet.write(18, 0, text)?;

    let data_validation = DataValidation::new().allow_time(DataValidationRule::Between(
        ExcelDateTime::parse_from_str("6:00")?,
        ExcelDateTime::parse_from_str("12:00")?,
    ));

    worksheet.add_data_validation(18, 1, 18, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 10. Limiting input to a string greater than a fixed length.
    // -----------------------------------------------------------------------
    let text = "Enter a string longer than 3 characters";
    worksheet.write(20, 0, text)?;

    let data_validation =
        DataValidation::new().allow_text_length(DataValidationRule::GreaterThan(3));

    worksheet.add_data_validation(20, 1, 20, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 11. Limiting input based on a formula.
    // -----------------------------------------------------------------------
    let text = "Enter a value if the following is true '=AND(F5=50,G5=60)'";
    worksheet.write(22, 0, text)?;

    let data_validation = DataValidation::new().allow_custom("=AND(F5=50,G5=60)".into());

    worksheet.add_data_validation(22, 1, 22, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 12. Displaying and modifying data validation messages.
    // -----------------------------------------------------------------------
    let text = "Displays a message when you select the cell";
    worksheet.write(24, 0, text)?;

    let data_validation = DataValidation::new()
        .allow_whole_number(DataValidationRule::Between(1, 100))
        .set_input_title("Enter an integer:")?
        .set_input_message("between 1 and 100")?;

    worksheet.add_data_validation(24, 1, 24, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 13. Displaying and modifying data validation messages.
    // -----------------------------------------------------------------------
    let text = "Display a custom error message when integer isn't between 1 and 100";
    worksheet.write(26, 0, text)?;

    let data_validation = DataValidation::new()
        .allow_whole_number(DataValidationRule::Between(1, 100))
        .set_input_title("Enter an integer:")?
        .set_input_message("between 1 and 100")?
        .set_error_title("Input value is not valid!")?
        .set_error_message("It should be an integer between 1 and 100")?;

    worksheet.add_data_validation(26, 1, 26, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 14. Displaying and modifying data validation messages.
    // -----------------------------------------------------------------------
    let text = "Display a custom info message when integer isn't between 1 and 100";
    worksheet.write(28, 0, text)?;

    let data_validation = DataValidation::new()
        .allow_whole_number(DataValidationRule::Between(1, 100))
        .set_input_title("Enter an integer:")?
        .set_input_message("between 1 and 100")?
        .set_error_title("Input value is not valid!")?
        .set_error_message("It should be an integer between 1 and 100")?
        .set_error_style(DataValidationErrorStyle::Information);

    worksheet.add_data_validation(28, 1, 28, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Save and close the file.
    // -----------------------------------------------------------------------
    workbook.save("data_validation.xlsx")?;

    Ok(())
}
