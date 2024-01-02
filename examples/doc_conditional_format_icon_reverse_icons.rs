// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of adding icon style conditional formatting to a worksheet. In the
//! second example the order of the icons is reversed.

use rust_xlsxwriter::{ConditionalFormatIconSet, ConditionalFormatIconType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write some captions.
    worksheet.write(1, 0, "Three Traffic lights - Green is highest")?;
    worksheet.write(2, 0, "Reversed - Red is highest")?;

    // Set the column width for clarity.
    worksheet.set_column_width(0, 35)?;

    // Write the worksheet data.
    worksheet.write_row(1, 1, [1, 2, 3])?;
    worksheet.write_row(2, 1, [1, 2, 3])?;

    // Three Traffic lights - Green is highest.
    let conditional_format = ConditionalFormatIconSet::new()
        .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights);

    worksheet.add_conditional_format(1, 1, 1, 3, &conditional_format)?;

    // Reversed - Red is highest.
    let conditional_format = ConditionalFormatIconSet::new()
        .set_icon_type(ConditionalFormatIconType::ThreeTrafficLights)
        .reverse_icons(true);

    worksheet.add_conditional_format(2, 1, 2, 3, &conditional_format)?;

    // Save the file.
    workbook.save("conditional_format.xlsx")?;

    Ok(())
}
