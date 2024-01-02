// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of adding icon style conditional formatting to a worksheet. In the
//! second example the default icons are changed.

use rust_xlsxwriter::{
    ConditionalFormatCustomIcon, ConditionalFormatIconSet, ConditionalFormatIconType,
    ConditionalFormatType, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write the worksheet data.
    worksheet.write_row(1, 1, [1, 2, 3])?;
    worksheet.write_row(2, 1, [1, 2, 3])?;

    // Three Traffic lights with default icons.
    let conditional_format = ConditionalFormatIconSet::new()
        .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights);

    worksheet.add_conditional_format(1, 1, 1, 3, &conditional_format)?;

    // Create some custom icons. Note, it is also required to set the default rules.
    let icons = [
        // We leave the default icon in the first/lowest position.
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
        ConditionalFormatCustomIcon::new()
            .set_rule(ConditionalFormatType::Percent, 33)
            .set_icon_type(ConditionalFormatIconType::FourHistograms, 0),
        ConditionalFormatCustomIcon::new()
            .set_rule(ConditionalFormatType::Percent, 67)
            .set_icon_type(ConditionalFormatIconType::FiveBoxes, 4),
    ];

    // Three Traffic lights with user defined icons.
    let conditional_format = ConditionalFormatIconSet::new()
        .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights)
        .set_icons(&icons);

    worksheet.add_conditional_format(2, 1, 2, 3, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
