// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of setting headers and footers in worksheets using the
//! rust_xlsxwriter library.

use rust_xlsxwriter::{HeaderImagePosition, Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // -----------------------------------------------------------------------
    // A simple example to start.
    // -----------------------------------------------------------------------
    let worksheet1 = workbook.add_worksheet().set_name("Simple")?;

    // Set page layout view so the headers/footers are visible.
    worksheet1.set_view_page_layout();

    // Add some sample text.
    worksheet1.write_string(0, 0, "Some text")?;

    worksheet1.set_header("&CHere is some centered text.");
    worksheet1.set_footer("&LHere is some left aligned text.");

    // -----------------------------------------------------------------------
    // This is an example of some of the header/footer variables.
    // -----------------------------------------------------------------------
    let worksheet2 = workbook.add_worksheet().set_name("Variables")?;
    worksheet2.set_view_page_layout();
    worksheet2.write_string(0, 0, "Some text")?;

    // Note the sections separators "&L" (left) "&C" (center) and "&R" (right).
    worksheet2.set_header("&LPage &[Page] of &[Pages]&CFilename: &[File]&RSheetname: &[Tab]");
    worksheet2.set_footer("&LCurrent date: &D&RCurrent time: &T");

    // -----------------------------------------------------------------------
    // This is an example of setting a header image.
    // -----------------------------------------------------------------------
    let worksheet3 = workbook.add_worksheet().set_name("Images")?;
    worksheet3.set_view_page_layout();
    worksheet3.write_string(0, 0, "Some text")?;

    let image = Image::new("examples/rust_logo.png")?
        .set_scale_height(0.5)
        .set_scale_width(0.5);

    worksheet3.set_header("&L&[Picture]");
    worksheet3.set_header_image(&image, HeaderImagePosition::Left)?;

    // Increase the top margin to 1.2 for clarity. The -1.0 values are ignored.
    worksheet3.set_margins(-1.0, -1.0, 1.2, -1.0, -1.0, -1.0);

    // -----------------------------------------------------------------------
    // This example shows how to use more than one font.
    // -----------------------------------------------------------------------
    let worksheet4 = workbook.add_worksheet().set_name("Mixed fonts")?;
    worksheet4.set_view_page_layout();
    worksheet4.write_string(0, 0, "Some text")?;

    worksheet4.set_header(r#"&C&"Courier New,Bold"Hello &"Arial,Italic"World"#);
    worksheet4.set_footer(r#"&C&"Symbol"e&"Arial" = mc&X2"#);

    // -----------------------------------------------------------------------
    // Example of line wrapping.
    // -----------------------------------------------------------------------
    let worksheet5 = workbook.add_worksheet().set_name("Word wrap")?;
    worksheet5.set_view_page_layout();
    worksheet5.write_string(0, 0, "Some text")?;

    worksheet5.set_header("&CHeading 1\nHeading 2");

    // -----------------------------------------------------------------------
    // Example of inserting a literal ampersand &.
    // -----------------------------------------------------------------------
    let worksheet6 = workbook.add_worksheet().set_name("Ampersand")?;
    worksheet6.set_view_page_layout();
    worksheet6.write_string(0, 0, "Some text")?;

    worksheet6.set_header("&CCuriouser && Curiouser - Attorneys at Law");

    workbook.save("headers_footers.xlsx")?;

    Ok(())
}
