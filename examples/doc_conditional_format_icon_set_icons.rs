// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! Example of adding icon style conditional formatting to a worksheet. In the
//! second example the default rules are changed.

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

    // Three Traffic lights with default rules.
    let conditional_format = ConditionalFormatIconSet::new()
        .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights);

    worksheet.add_conditional_format(1, 1, 1, 3, &conditional_format)?;

    // Create some user defined rules. The number of rules (3-5) must match the
    // number of symbols in the icon type and the first rule should always be
    // "0%" (this is the default in Excel).
    let icons = [
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percent, 0),
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percentile, 50),
        ConditionalFormatCustomIcon::new().set_rule(ConditionalFormatType::Percentile, 90),
    ];

    // Three Traffic lights with user defined rules.
    let conditional_format = ConditionalFormatIconSet::new()
        .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights)
        .set_icons(&icons);

    worksheet.add_conditional_format(2, 1, 2, 3, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
