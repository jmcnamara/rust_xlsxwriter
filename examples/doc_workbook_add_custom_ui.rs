// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! An example of adding a custom ribbon tab to a workbook using a Custom UI XML
//! file. The tab contains a button that runs a macro from a VBA project.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add the Custom UI XML file for the Excel 2010+ ribbon.
    workbook.add_custom_ui("examples/customUI14.xml")?;

    // Add the VBA project with the macro used by the ribbon button.
    workbook.add_vba_project("examples/vbaProject_custom_ui.bin")?;

    let worksheet = workbook.add_worksheet();
    worksheet.write(0, 0, "Use the 'My Ribbon' tab.")?;

    // Note the `.xlsm` extension, which is required for VBA macros.
    workbook.save("custom_ui.xlsm")?;

    Ok(())
}
