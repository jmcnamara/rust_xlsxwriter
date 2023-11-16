// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! Example of how to add conditional formatting to a worksheet using the
//! rust_xlsxwriter library.
//!
//! Conditional formatting allows you to apply a format to a cell or a range of
//! cells based on certain criteria.

use rust_xlsxwriter::{
    ConditionalFormat2ColorScale, ConditionalFormat3ColorScale, ConditionalFormatAverage,
    ConditionalFormatAverageCriteria, ConditionalFormatCell, ConditionalFormatCellCriteria,
    ConditionalFormatDataBar, ConditionalFormatDataBarDirection, ConditionalFormatDuplicate,
    ConditionalFormatText, ConditionalFormatTextCriteria, ConditionalFormatTop, Format, Workbook,
    XlsxError,
};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a format. Light red fill with dark red text.
    let format1 = Format::new()
        .set_font_color("9C0006")
        .set_background_color("FFC7CE");

    // Add a format. Green fill with dark green text.
    let format2 = Format::new()
        .set_font_color("006100")
        .set_background_color("C6EFCE");

    // some sample data to run the conditional formatting against.
    let data = [
        [34, 72, 38, 30, 75, 48, 75, 66, 84, 86],
        [6, 24, 1, 84, 54, 62, 60, 3, 26, 59],
        [28, 79, 97, 13, 85, 93, 93, 22, 5, 14],
        [27, 71, 40, 17, 18, 79, 90, 93, 29, 47],
        [88, 25, 33, 23, 67, 1, 59, 79, 47, 36],
        [24, 100, 20, 88, 29, 33, 38, 54, 54, 88],
        [6, 57, 88, 28, 10, 26, 37, 7, 41, 48],
        [52, 78, 1, 96, 26, 45, 47, 33, 96, 36],
        [60, 54, 81, 66, 81, 90, 80, 93, 12, 55],
        [70, 5, 46, 14, 71, 19, 66, 36, 41, 21],
    ];

    // -----------------------------------------------------------------------
    // Example 1. Cell conditional formatting.
    // -----------------------------------------------------------------------
    let caption = "Cells with values >= 50 are in light red. Values < 50 are in light green.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    worksheet.write_row_matrix(2, 1, data)?;

    // Set the column widths for clarity.
    for col_num in 1..=10u16 {
        worksheet.set_column_width(col_num, 6)?;
    }

    // Write a conditional format over a range.
    let conditional_format = ConditionalFormatCell::new()
        .set_criteria(ConditionalFormatCellCriteria::GreaterThanOrEqualTo)
        .set_value(50)
        .set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Write another conditional format over the same range.
    let conditional_format = ConditionalFormatCell::new()
        .set_criteria(ConditionalFormatCellCriteria::LessThan)
        .set_value(50)
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Example 2. Cell conditional formatting with between ranges.
    // -----------------------------------------------------------------------
    let caption =
        "Values between 30 and 70 are in light red. Values outside that range are in light green.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    worksheet.write_row_matrix(2, 1, data)?;

    // Set the column widths for clarity.
    for col_num in 1..=10u16 {
        worksheet.set_column_width(col_num, 6)?;
    }

    // Write a conditional format over a range.
    let conditional_format = ConditionalFormatCell::new()
        .set_criteria(ConditionalFormatCellCriteria::Between)
        .set_minimum(30)
        .set_maximum(70)
        .set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Write another conditional format over the same range.
    let conditional_format = ConditionalFormatCell::new()
        .set_criteria(ConditionalFormatCellCriteria::NotBetween)
        .set_minimum(30)
        .set_maximum(70)
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Example 3. Duplicate and Unique conditional formats.
    // -----------------------------------------------------------------------
    let caption = "Duplicate values are in light red. Unique values are in light green.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    worksheet.write_row_matrix(2, 1, data)?;

    // Set the column widths for clarity.
    for col_num in 1..=10u16 {
        worksheet.set_column_width(col_num, 6)?;
    }

    // Write a conditional format over a range.
    let conditional_format = ConditionalFormatDuplicate::new().set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Invert the duplicate conditional format to show unique values in the
    // same range.
    let conditional_format = ConditionalFormatDuplicate::new()
        .invert()
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Example 4. Above and Below Average conditional formats.
    // -----------------------------------------------------------------------
    let caption = "Above average values are in light red. Below average values are in light green.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    worksheet.write_row_matrix(2, 1, data)?;

    // Set the column widths for clarity.
    for col_num in 1..=10u16 {
        worksheet.set_column_width(col_num, 6)?;
    }

    // Write a conditional format over a range. The default criteria is Above Average.
    let conditional_format = ConditionalFormatAverage::new().set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Write another conditional format over the same range.
    let conditional_format = ConditionalFormatAverage::new()
        .set_criteria(ConditionalFormatAverageCriteria::BelowAverage)
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Example 5. Top and Bottom range conditional formats.
    // -----------------------------------------------------------------------
    let caption = "Top 10 values are in light red. Bottom 10 values are in light green.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    worksheet.write_row_matrix(2, 1, data)?;

    // Set the column widths for clarity.
    for col_num in 1..=10u16 {
        worksheet.set_column_width(col_num, 6)?;
    }

    // Write a conditional format over a range.
    let conditional_format = ConditionalFormatTop::new()
        .set_value(10)
        .set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Invert the Top conditional format to show Bottom values in the same
    // range.
    let conditional_format = ConditionalFormatTop::new()
        .invert()
        .set_value(10)
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Example 6. Cell conditional formatting in non-contiguous range.
    // -----------------------------------------------------------------------
    let caption = "Cells with values >= 50 are in light red. Values < 50 are in light green. Non-contiguous ranges.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    worksheet.write_row_matrix(2, 1, data)?;

    // Set the column widths for clarity.
    for col_num in 1..=10u16 {
        worksheet.set_column_width(col_num, 6)?;
    }

    // Write a conditional format over a non-contiguous range.
    let conditional_format = ConditionalFormatCell::new()
        .set_criteria(ConditionalFormatCellCriteria::GreaterThanOrEqualTo)
        .set_value(50)
        .set_multi_range("B3:D6 I3:K6 B9:D12 I9:K12")
        .set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Write another conditional format over the same range.
    let conditional_format = ConditionalFormatCell::new()
        .set_criteria(ConditionalFormatCellCriteria::LessThan)
        .set_value(50)
        .set_multi_range("B3:D6 I3:K6 B9:D12 I9:K12")
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Example 7. Text style conditional formats.
    // -----------------------------------------------------------------------
    let caption =
        "Column A shows words that contain the sub-word 'rust'. Column C shows words that start/end with 't'";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 0, caption)?;

    // Add some sample data.
    let word_list = [
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
    worksheet.write_column(1, 0, word_list)?;
    worksheet.write_column(1, 2, word_list)?;

    // Set the column widths for clarity.
    worksheet.set_column_width(0, 20)?;
    worksheet.set_column_width(2, 20)?;

    // Write a text "containing" conditional format over a range.
    let conditional_format = ConditionalFormatText::new()
        .set_criteria(ConditionalFormatTextCriteria::Contains)
        .set_value("rust")
        .set_format(&format2);

    worksheet.add_conditional_format(1, 0, 13, 0, &conditional_format)?;

    // Write a text "not containing" conditional format over the same range.
    let conditional_format = ConditionalFormatText::new()
        .set_criteria(ConditionalFormatTextCriteria::DoesNotContain)
        .set_value("rust")
        .set_format(&format1);

    worksheet.add_conditional_format(1, 0, 13, 0, &conditional_format)?;

    // Write a text "begins with" conditional format over a range.
    let conditional_format = ConditionalFormatText::new()
        .set_criteria(ConditionalFormatTextCriteria::BeginsWith)
        .set_value("t")
        .set_format(&format2);

    worksheet.add_conditional_format(1, 2, 13, 2, &conditional_format)?;

    // Write a text "ends with" conditional format over the same range.
    let conditional_format = ConditionalFormatText::new()
        .set_criteria(ConditionalFormatTextCriteria::EndsWith)
        .set_value("t")
        .set_format(&format1);

    worksheet.add_conditional_format(1, 2, 13, 2, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Example 8. Examples of 2 color scale conditional formats.
    // -----------------------------------------------------------------------
    let caption = "Examples of 2 color scale conditional formats";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    let data = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    worksheet.write_column(2, 1, data)?;
    worksheet.write_column(2, 3, data)?;
    worksheet.write_column(2, 5, data)?;
    worksheet.write_column(2, 7, data)?;
    worksheet.write_column(2, 9, data)?;
    worksheet.write_column(2, 11, data)?;

    // Set the column widths for clarity.
    for col_num in 0..=12u16 {
        worksheet.set_column_width(col_num, 6)?;
    }

    // Write 2 color scale formats with standard Excel colors.
    let conditional_format = ConditionalFormat2ColorScale::new()
        .set_minimum_color("F8696B")
        .set_maximum_color("FCFCFF");

    worksheet.add_conditional_format(2, 1, 11, 1, &conditional_format)?;

    let conditional_format = ConditionalFormat2ColorScale::new()
        .set_minimum_color("FCFCFF")
        .set_maximum_color("F8696B");

    worksheet.add_conditional_format(2, 3, 11, 3, &conditional_format)?;

    let conditional_format = ConditionalFormat2ColorScale::new()
        .set_minimum_color("FCFCFF")
        .set_maximum_color("63BE7B");

    worksheet.add_conditional_format(2, 5, 11, 5, &conditional_format)?;

    let conditional_format = ConditionalFormat2ColorScale::new()
        .set_minimum_color("63BE7B")
        .set_maximum_color("FCFCFF");

    worksheet.add_conditional_format(2, 7, 11, 7, &conditional_format)?;

    let conditional_format = ConditionalFormat2ColorScale::new()
        .set_minimum_color("FFEF9C")
        .set_maximum_color("63BE7B");

    worksheet.add_conditional_format(2, 9, 11, 9, &conditional_format)?;

    let conditional_format = ConditionalFormat2ColorScale::new()
        .set_minimum_color("63BE7B")
        .set_maximum_color("FFEF9C");

    worksheet.add_conditional_format(2, 11, 11, 11, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Example 9. Examples of 3 color scale conditional formats.
    // -----------------------------------------------------------------------
    let caption = "Examples of 3 color scale conditional formats";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    let data = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    worksheet.write_column(2, 1, data)?;
    worksheet.write_column(2, 3, data)?;
    worksheet.write_column(2, 5, data)?;
    worksheet.write_column(2, 7, data)?;
    worksheet.write_column(2, 9, data)?;
    worksheet.write_column(2, 11, data)?;

    // Set the column widths for clarity.
    for col_num in 0..=12u16 {
        worksheet.set_column_width(col_num, 6)?;
    }

    // Write 3 color scale formats with standard Excel colors.
    let conditional_format = ConditionalFormat3ColorScale::new()
        .set_minimum_color("F8696B")
        .set_midpoint_color("FFEB84")
        .set_maximum_color("63BE7B");

    worksheet.add_conditional_format(2, 1, 11, 1, &conditional_format)?;

    let conditional_format = ConditionalFormat3ColorScale::new()
        .set_minimum_color("63BE7B")
        .set_midpoint_color("FFEB84")
        .set_maximum_color("F8696B");

    worksheet.add_conditional_format(2, 3, 11, 3, &conditional_format)?;

    let conditional_format = ConditionalFormat3ColorScale::new()
        .set_minimum_color("F8696B")
        .set_midpoint_color("FCFCFF")
        .set_maximum_color("63BE7B");

    worksheet.add_conditional_format(2, 5, 11, 5, &conditional_format)?;

    let conditional_format = ConditionalFormat3ColorScale::new()
        .set_minimum_color("63BE7B")
        .set_midpoint_color("FCFCFF")
        .set_maximum_color("F8696B");

    worksheet.add_conditional_format(2, 7, 11, 7, &conditional_format)?;

    let conditional_format = ConditionalFormat3ColorScale::new()
        .set_minimum_color("F8696B")
        .set_midpoint_color("FCFCFF")
        .set_maximum_color("5A8AC6");

    worksheet.add_conditional_format(2, 9, 11, 9, &conditional_format)?;

    let conditional_format = ConditionalFormat3ColorScale::new()
        .set_minimum_color("5A8AC6")
        .set_midpoint_color("FCFCFF")
        .set_maximum_color("F8696B");

    worksheet.add_conditional_format(2, 11, 11, 11, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Example 9. Examples of data bars.
    // -----------------------------------------------------------------------
    let caption = "Examples of data bars";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;
    worksheet.write(1, 1, "Default")?;
    worksheet.write(1, 3, "Default negative")?;
    worksheet.write(1, 5, "User color")?;
    worksheet.write(1, 7, "Changed direction")?;

    // Write the worksheet data.
    let data1 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    let data2 = [6, 4, 2, -2, -4, -6, -4, -2, 2, 4];
    worksheet.write_column(2, 1, data1)?;
    worksheet.write_column(2, 3, data2)?;
    worksheet.write_column(2, 5, data1)?;
    worksheet.write_column(2, 7, data1)?;

    // Write a standard Excel data bar.
    let conditional_format = ConditionalFormatDataBar::new();

    worksheet.add_conditional_format(2, 1, 11, 1, &conditional_format)?;

    // Write a standard Excel data bar with negative data
    let conditional_format = ConditionalFormatDataBar::new();

    worksheet.add_conditional_format(2, 3, 11, 3, &conditional_format)?;

    // Write a data bar with a user defined fill color.
    let conditional_format = ConditionalFormatDataBar::new().set_fill_color("009933");

    worksheet.add_conditional_format(2, 5, 11, 5, &conditional_format)?;

    // Write a data bar with the direction changed.
    let conditional_format = ConditionalFormatDataBar::new()
        .set_direction(ConditionalFormatDataBarDirection::RightToLeft);

    worksheet.add_conditional_format(2, 7, 11, 7, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Save and close the file.
    // -----------------------------------------------------------------------
    workbook.save("conditional_formats.xlsx")?;

    Ok(())
}
