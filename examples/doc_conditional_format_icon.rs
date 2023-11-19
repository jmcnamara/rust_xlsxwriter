// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! Example of adding icon style conditional formatting to a worksheet.

use rust_xlsxwriter::{
    ConditionalFormatIconSet, ConditionalFormatIconType, Format, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add a format for headers.
    let bold = Format::new().set_bold();

    // Add a format for descriptions.
    let indent = Format::new().set_indent(2);

    // Write some captions.
    worksheet.write_with_format(0, 0, "Icon style conditional formats", &bold)?;
    worksheet.write_with_format(1, 0, "Three Traffic lights - Green is highest", &indent)?;
    worksheet.write_with_format(2, 0, "Reversed - Red is highest", &indent)?;
    worksheet.write_with_format(3, 0, "Icons only - The number data is hidden", &indent)?;

    worksheet.write_with_format(4, 0, "Other three-five icon examples", &bold)?;
    worksheet.write_with_format(5, 0, "Three arrows", &indent)?;
    worksheet.write_with_format(6, 0, "Three symbols", &indent)?;
    worksheet.write_with_format(7, 0, "Three flags", &indent)?;

    worksheet.write_with_format(8, 0, "Four arrows", &indent)?;
    worksheet.write_with_format(9, 0, "Four circles - Red (highest) to Black", &indent)?;
    worksheet.write_with_format(10, 0, "Four rating histograms", &indent)?;

    worksheet.write_with_format(11, 0, "Five arrows", &indent)?;
    worksheet.write_with_format(12, 0, "Five rating histograms", &indent)?;
    worksheet.write_with_format(13, 0, "Five rating quadrants", &indent)?;

    // Set the column width for clarity.
    worksheet.set_column_width(0, 35)?;

    // Write the worksheet data.
    let data3 = [1, 2, 3];
    let data4 = [1, 2, 3, 4];
    let data5 = [1, 2, 3, 4, 5];
    worksheet.write_row(1, 1, data3)?;
    worksheet.write_row(2, 1, data3)?;
    worksheet.write_row(3, 1, data3)?;

    worksheet.write_row(5, 1, data3)?;
    worksheet.write_row(6, 1, data3)?;
    worksheet.write_row(7, 1, data3)?;

    worksheet.write_row(8, 1, data4)?;
    worksheet.write_row(9, 1, data4)?;
    worksheet.write_row(10, 1, data4)?;

    worksheet.write_row(11, 1, data5)?;
    worksheet.write_row(12, 1, data5)?;
    worksheet.write_row(13, 1, data5)?;

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

    // Three flags.
    let conditional_format =
        ConditionalFormatIconSet::new().set_icon_type(ConditionalFormatIconType::ThreeFlags);

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

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
