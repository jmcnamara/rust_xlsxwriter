// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of how to add conditional formatting to a worksheet using the
//! rust_xlsxwriter library.
//!
//! Conditional formatting allows you to apply a format to a cell or a range of
//! cells based on user defined rule.

use rust_xlsxwriter::{
    ConditionalFormat2ColorScale, ConditionalFormat3ColorScale, ConditionalFormatAverage,
    ConditionalFormatAverageRule, ConditionalFormatCell, ConditionalFormatCellRule,
    ConditionalFormatDataBar, ConditionalFormatDataBarDirection, ConditionalFormatDuplicate,
    ConditionalFormatFormula, ConditionalFormatIconSet, ConditionalFormatIconType,
    ConditionalFormatText, ConditionalFormatTextRule, ConditionalFormatTop, Format, Workbook,
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

    // Add a format for headers.
    let bold = Format::new().set_bold();

    // Add a format for descriptions.
    let indent = Format::new().set_indent(2);

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
    // Worksheet 1. Cell conditional formatting.
    // -----------------------------------------------------------------------
    let caption = "Cells with values >= 50 are in light red. Values < 50 are in light green.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    worksheet.write_row_matrix(2, 1, data)?;

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 10, 6)?;

    // Write a conditional format over a range.
    let conditional_format = ConditionalFormatCell::new()
        .set_rule(ConditionalFormatCellRule::GreaterThanOrEqualTo(50))
        .set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Write another conditional format over the same range.
    let conditional_format = ConditionalFormatCell::new()
        .set_rule(ConditionalFormatCellRule::LessThan(50))
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Worksheet 2. Cell conditional formatting with between ranges.
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
    worksheet.set_column_range_width(1, 10, 6)?;

    // Write a conditional format over a range.
    let conditional_format = ConditionalFormatCell::new()
        .set_rule(ConditionalFormatCellRule::Between(30, 70))
        .set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Write another conditional format over the same range.
    let conditional_format = ConditionalFormatCell::new()
        .set_rule(ConditionalFormatCellRule::NotBetween(30, 70))
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Worksheet 3. Duplicate and Unique conditional formats.
    // -----------------------------------------------------------------------
    let caption = "Duplicate values are in light red. Unique values are in light green.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    worksheet.write_row_matrix(2, 1, data)?;

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 10, 6)?;

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
    // Worksheet 4. Above and Below Average conditional formats.
    // -----------------------------------------------------------------------
    let caption = "Above average values are in light red. Below average values are in light green.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    worksheet.write_row_matrix(2, 1, data)?;

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 10, 6)?;

    // Write a conditional format over a range. The default criteria is Above Average.
    let conditional_format = ConditionalFormatAverage::new().set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Write another conditional format over the same range.
    let conditional_format = ConditionalFormatAverage::new()
        .set_rule(ConditionalFormatAverageRule::BelowAverage)
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Worksheet 5. Top and Bottom range conditional formats.
    // -----------------------------------------------------------------------
    let caption = "Top 10 values are in light red. Bottom 10 values are in light green.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    worksheet.write_row_matrix(2, 1, data)?;

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 10, 6)?;

    // Write a conditional format over a range.
    let conditional_format = ConditionalFormatTop::new()
        .set_rule(rust_xlsxwriter::ConditionalFormatTopRule::Top(10))
        .set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Also show the bottom values in the same range.
    let conditional_format = ConditionalFormatTop::new()
        .set_rule(rust_xlsxwriter::ConditionalFormatTopRule::Bottom(10))
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Worksheet 6. Cell conditional formatting in non-contiguous range.
    // -----------------------------------------------------------------------
    let caption = "Cells with values >= 50 are in light red. Values < 50 are in light green. Non-contiguous ranges.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    worksheet.write_row_matrix(2, 1, data)?;

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 10, 6)?;

    // Write a conditional format over a non-contiguous range.
    let conditional_format = ConditionalFormatCell::new()
        .set_rule(ConditionalFormatCellRule::GreaterThanOrEqualTo(50))
        .set_multi_range("B3:D6 I3:K6 B9:D12 I9:K12")
        .set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Write another conditional format over the same range.
    let conditional_format = ConditionalFormatCell::new()
        .set_rule(ConditionalFormatCellRule::LessThan(50))
        .set_multi_range("B3:D6 I3:K6 B9:D12 I9:K12")
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Worksheet 7. Formula conditional formatting.
    // -----------------------------------------------------------------------
    let caption = "Even numbered cells are in light green. Odd numbered cells are in light red.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the worksheet data.
    worksheet.write_row_matrix(2, 1, data)?;

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 10, 6)?;

    // Write a conditional format over a range.
    let conditional_format = ConditionalFormatFormula::new()
        .set_rule("=ISODD(B3)")
        .set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Write another conditional format over the same range.
    let conditional_format = ConditionalFormatFormula::new()
        .set_rule("=ISEVEN(B3)")
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Worksheet 8. Text style conditional formats.
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
        .set_rule(ConditionalFormatTextRule::Contains("rust".to_string()))
        .set_format(&format2);

    worksheet.add_conditional_format(1, 0, 13, 0, &conditional_format)?;

    // Write a text "not containing" conditional format over the same range.
    let conditional_format = ConditionalFormatText::new()
        .set_rule(ConditionalFormatTextRule::DoesNotContain(
            "rust".to_string(),
        ))
        .set_format(&format1);

    worksheet.add_conditional_format(1, 0, 13, 0, &conditional_format)?;

    // Write a text "begins with" conditional format over a range.
    let conditional_format = ConditionalFormatText::new()
        .set_rule(ConditionalFormatTextRule::BeginsWith("t".to_string()))
        .set_format(&format2);

    worksheet.add_conditional_format(1, 2, 13, 2, &conditional_format)?;

    // Write a text "ends with" conditional format over the same range.
    let conditional_format = ConditionalFormatText::new()
        .set_rule(ConditionalFormatTextRule::EndsWith("t".to_string()))
        .set_format(&format1);

    worksheet.add_conditional_format(1, 2, 13, 2, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Worksheet 9. Examples of 2 color scale conditional formats.
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
    worksheet.set_column_range_width(0, 12, 6)?;

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
    // Worksheet 10. Examples of 3 color scale conditional formats.
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
    worksheet.set_column_range_width(0, 12, 6)?;

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
    // Worksheet 11. Examples of data bars.
    // -----------------------------------------------------------------------
    let caption = "Examples of data bars";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write_with_format(0, 1, caption, &bold)?;
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
    // Worksheet 12. Examples of icon style conditional formats.
    // -----------------------------------------------------------------------
    let caption = "Examples of icon style conditional formats.";

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the caption.
    worksheet.write_with_format(0, 0, caption, &bold)?;
    worksheet.write_with_format(1, 0, "Three Traffic lights - Green is highest", &indent)?;
    worksheet.write_with_format(2, 0, "Reversed - Red is highest", &indent)?;
    worksheet.write_with_format(3, 0, "Icons only - The number data is hidden", &indent)?;

    worksheet.write_with_format(4, 0, "Other three-five icon examples", &bold)?;
    worksheet.write_with_format(5, 0, "Three arrows", &indent)?;
    worksheet.write_with_format(6, 0, "Three symbols", &indent)?;
    worksheet.write_with_format(7, 0, "Three stars", &indent)?;

    worksheet.write_with_format(8, 0, "Four arrows", &indent)?;
    worksheet.write_with_format(9, 0, "Four circles - Red (highest) to Black", &indent)?;
    worksheet.write_with_format(10, 0, "Four rating histograms", &indent)?;

    worksheet.write_with_format(11, 0, "Five arrows", &indent)?;
    worksheet.write_with_format(12, 0, "Five rating histograms", &indent)?;
    worksheet.write_with_format(13, 0, "Five rating quadrants", &indent)?;

    // Set the column width for clarity.
    worksheet.set_column_width(0, 35)?;

    // Write the worksheet data.
    worksheet.write_row(1, 1, [1, 2, 3])?;
    worksheet.write_row(2, 1, [1, 2, 3])?;
    worksheet.write_row(3, 1, [1, 2, 3])?;

    worksheet.write_row(5, 1, [1, 2, 3])?;
    worksheet.write_row(6, 1, [1, 2, 3])?;
    worksheet.write_row(7, 1, [1, 2, 3])?;

    worksheet.write_row(8, 1, [1, 2, 3, 4])?;
    worksheet.write_row(9, 1, [1, 2, 3, 4])?;
    worksheet.write_row(10, 1, [1, 2, 3, 4])?;

    worksheet.write_row(11, 1, [1, 2, 3, 4, 5])?;
    worksheet.write_row(12, 1, [1, 2, 3, 4, 5])?;
    worksheet.write_row(13, 1, [1, 2, 3, 4, 5])?;

    // Three Traffic lights - Green is highest.
    let conditional_format = ConditionalFormatIconSet::new()
        .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights);

    worksheet.add_conditional_format(1, 1, 1, 3, &conditional_format)?;

    // Reversed - Red is highest.
    let conditional_format = ConditionalFormatIconSet::new()
        .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights)
        .reverse_icons(true);

    worksheet.add_conditional_format(2, 1, 2, 3, &conditional_format)?;

    // Icons only - The number data is hidden.
    let conditional_format = ConditionalFormatIconSet::new()
        .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights)
        .show_icons_only(true);

    worksheet.add_conditional_format(3, 1, 3, 3, &conditional_format)?;

    // Three arrows.
    let conditional_format =
        ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::ThreeArrows);

    worksheet.add_conditional_format(5, 1, 5, 3, &conditional_format)?;

    // Three symbols.
    let conditional_format = ConditionalFormatIconSet::new()
        .set_icon_type(ConditionalFormatIconType::ThreeSymbolsCircled);

    worksheet.add_conditional_format(6, 1, 6, 3, &conditional_format)?;

    // Three stars.
    let conditional_format =
        ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::ThreeStars);

    worksheet.add_conditional_format(7, 1, 7, 3, &conditional_format)?;

    // Four Arrows.
    let conditional_format =
        ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::FourArrows);

    worksheet.add_conditional_format(8, 1, 8, 4, &conditional_format)?;

    // Four circles - Red (highest) to Black (lowest).
    let conditional_format =
        ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::FourRedToBlack);

    worksheet.add_conditional_format(9, 1, 9, 4, &conditional_format)?;

    // Four rating histograms.
    let conditional_format =
        ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::FourHistograms);

    worksheet.add_conditional_format(10, 1, 10, 4, &conditional_format)?;

    // Four Arrows.
    let conditional_format =
        ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::FiveArrows);

    worksheet.add_conditional_format(11, 1, 11, 5, &conditional_format)?;

    // Four rating histograms.
    let conditional_format =
        ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::FiveHistograms);

    worksheet.add_conditional_format(12, 1, 12, 5, &conditional_format)?;

    // Four rating quadrants.
    let conditional_format =
        ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::FiveQuadrants);

    worksheet.add_conditional_format(13, 1, 13, 5, &conditional_format)?;

    // -----------------------------------------------------------------------
    // Save and close the file.
    // -----------------------------------------------------------------------
    workbook.save("conditional_formats.xlsx")?;

    Ok(())
}
