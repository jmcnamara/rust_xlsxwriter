// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of how to use the rust_xlsxwriter library to write formulas and
//! functions that create dynamic arrays. These functions are new to Excel
//! 365. The examples mirror the examples in the Excel documentation for these
//! functions.
use rust_xlsxwriter::{Color, Format, Workbook, Worksheet, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create some header formats to use in the worksheets.
    let header1 = Format::new()
        .set_foreground_color(Color::RGB(0x74AC4C))
        .set_font_color(Color::RGB(0xFFFFFF));

    let header2 = Format::new()
        .set_foreground_color(Color::RGB(0x528FD3))
        .set_font_color(Color::RGB(0xFFFFFF));

    // -----------------------------------------------------------------------
    // Example of using the FILTER() function.
    // -----------------------------------------------------------------------
    let worksheet1 = workbook.add_worksheet().set_name("Filter")?;

    worksheet1.write_dynamic_formula(1, 5, "=FILTER(A1:D17,C1:C17=K2)")?;

    // Write the data the function will work on.
    worksheet1.write_string_with_format(0, 10, "Product", &header2)?;
    worksheet1.write_string(1, 10, "Apple")?;
    worksheet1.write_string_with_format(0, 5, "Region", &header2)?;
    worksheet1.write_string_with_format(0, 6, "Sales Rep", &header2)?;
    worksheet1.write_string_with_format(0, 7, "Product", &header2)?;
    worksheet1.write_string_with_format(0, 8, "Units", &header2)?;

    // Add sample worksheet data to work on.
    write_worksheet_data(worksheet1, &header1)?;
    worksheet1.set_column_width_pixels(4, 20)?;
    worksheet1.set_column_width_pixels(9, 20)?;

    // -----------------------------------------------------------------------
    // Example of using the UNIQUE() function.
    // -----------------------------------------------------------------------
    let worksheet2 = workbook.add_worksheet().set_name("Unique")?;

    worksheet2.write_dynamic_formula(1, 5, "=UNIQUE(B2:B17)")?;

    // A more complex example combining SORT and UNIQUE.
    worksheet2.write_dynamic_formula(1, 7, "SORT(UNIQUE(B2:B17))")?;

    // Write the data the function will work on.
    worksheet2.write_string_with_format(0, 5, "Sales Rep", &header2)?;
    worksheet2.write_string_with_format(0, 7, "Sales Rep", &header2)?;

    // Add sample worksheet data to work on.
    write_worksheet_data(worksheet2, &header1)?;
    worksheet2.set_column_width_pixels(4, 20)?;
    worksheet2.set_column_width_pixels(6, 20)?;

    // -----------------------------------------------------------------------
    // Example of using the SORT() function.
    // -----------------------------------------------------------------------
    let worksheet3 = workbook.add_worksheet().set_name("Sort")?;

    // A simple SORT example.
    worksheet3.write_dynamic_formula(1, 5, "=SORT(B2:B17)")?;

    // A more complex example combining SORT and FILTER.
    worksheet3.write_dynamic_formula(1, 7, r#"=SORT(FILTER(C2:D17,D2:D17>5000,""),2,1)"#)?;

    // Write the data the function will work on.
    worksheet3.write_string_with_format(0, 5, "Sales Rep", &header2)?;
    worksheet3.write_string_with_format(0, 7, "Product", &header2)?;
    worksheet3.write_string_with_format(0, 8, "Units", &header2)?;

    // Add sample worksheet data to work on.
    write_worksheet_data(worksheet3, &header1)?;
    worksheet3.set_column_width_pixels(4, 20)?;
    worksheet3.set_column_width_pixels(6, 20)?;

    // -----------------------------------------------------------------------
    // Example of using the SORTBY() function.
    // -----------------------------------------------------------------------
    let worksheet4 = workbook.add_worksheet().set_name("Sortby")?;

    worksheet4.write_dynamic_formula(1, 3, "=SORTBY(A2:B9,B2:B9)")?;

    // Write the data the function will work on.
    worksheet4.write_string_with_format(0, 0, "Name", &header1)?;
    worksheet4.write_string_with_format(0, 1, "Age", &header1)?;

    worksheet4.write_string(1, 0, "Tom")?;
    worksheet4.write_string(2, 0, "Fred")?;
    worksheet4.write_string(3, 0, "Amy")?;
    worksheet4.write_string(4, 0, "Sal")?;
    worksheet4.write_string(5, 0, "Fritz")?;
    worksheet4.write_string(6, 0, "Srivan")?;
    worksheet4.write_string(7, 0, "Xi")?;
    worksheet4.write_string(8, 0, "Hector")?;

    worksheet4.write_number(1, 1, 52)?;
    worksheet4.write_number(2, 1, 65)?;
    worksheet4.write_number(3, 1, 22)?;
    worksheet4.write_number(4, 1, 73)?;
    worksheet4.write_number(5, 1, 19)?;
    worksheet4.write_number(6, 1, 39)?;
    worksheet4.write_number(7, 1, 19)?;
    worksheet4.write_number(8, 1, 66)?;

    worksheet4.write_string_with_format(0, 3, "Name", &header2)?;
    worksheet4.write_string_with_format(0, 4, "Age", &header2)?;

    worksheet4.set_column_width_pixels(2, 20)?;

    // -----------------------------------------------------------------------
    // Example of using the XLOOKUP() function.
    // -----------------------------------------------------------------------
    let worksheet5 = workbook.add_worksheet().set_name("Xlookup")?;

    worksheet5.write_dynamic_formula(0, 5, "=XLOOKUP(E1,A2:A9,C2:C9)")?;

    // Write the data the function will work on.
    worksheet5.write_string_with_format(0, 0, "Country", &header1)?;
    worksheet5.write_string_with_format(0, 1, "Abr", &header1)?;
    worksheet5.write_string_with_format(0, 2, "Prefix", &header1)?;

    worksheet5.write_string(1, 0, "China")?;
    worksheet5.write_string(2, 0, "India")?;
    worksheet5.write_string(3, 0, "United States")?;
    worksheet5.write_string(4, 0, "Indonesia")?;
    worksheet5.write_string(5, 0, "Brazil")?;
    worksheet5.write_string(6, 0, "Pakistan")?;
    worksheet5.write_string(7, 0, "Nigeria")?;
    worksheet5.write_string(8, 0, "Bangladesh")?;

    worksheet5.write_string(1, 1, "CN")?;
    worksheet5.write_string(2, 1, "IN")?;
    worksheet5.write_string(3, 1, "US")?;
    worksheet5.write_string(4, 1, "ID")?;
    worksheet5.write_string(5, 1, "BR")?;
    worksheet5.write_string(6, 1, "PK")?;
    worksheet5.write_string(7, 1, "NG")?;
    worksheet5.write_string(8, 1, "BD")?;

    worksheet5.write_number(1, 2, 86)?;
    worksheet5.write_number(2, 2, 91)?;
    worksheet5.write_number(3, 2, 1)?;
    worksheet5.write_number(4, 2, 62)?;
    worksheet5.write_number(5, 2, 55)?;
    worksheet5.write_number(6, 2, 92)?;
    worksheet5.write_number(7, 2, 234)?;
    worksheet5.write_number(8, 2, 880)?;

    worksheet5.write_string_with_format(0, 4, "Brazil", &header2)?;

    worksheet5.set_column_width_pixels(0, 100)?;
    worksheet5.set_column_width_pixels(3, 20)?;

    // -----------------------------------------------------------------------
    // Example of using the XMATCH() function.
    // -----------------------------------------------------------------------
    let worksheet6 = workbook.add_worksheet().set_name("Xmatch")?;

    worksheet6.write_dynamic_formula(1, 3, "=XMATCH(C2,A2:A6)")?;

    // Write the data the function will work on.
    worksheet6.write_string_with_format(0, 0, "Product", &header1)?;

    worksheet6.write_string(1, 0, "Apple")?;
    worksheet6.write_string(2, 0, "Grape")?;
    worksheet6.write_string(3, 0, "Pear")?;
    worksheet6.write_string(4, 0, "Banana")?;
    worksheet6.write_string(5, 0, "Cherry")?;

    worksheet6.write_string_with_format(0, 2, "Product", &header2)?;
    worksheet6.write_string_with_format(0, 3, "Position", &header2)?;
    worksheet6.write_string(1, 2, "Grape")?;

    worksheet6.set_column_width_pixels(1, 20)?;

    // -----------------------------------------------------------------------
    // Example of using the RANDARRAY() function.
    // -----------------------------------------------------------------------
    let worksheet7 = workbook.add_worksheet().set_name("Randarray")?;

    worksheet7.write_dynamic_formula(0, 0, "=RANDARRAY(5,3,1,100, TRUE)")?;

    // -----------------------------------------------------------------------
    // Example of using the SEQUENCE() function.
    // -----------------------------------------------------------------------
    let worksheet8 = workbook.add_worksheet().set_name("Sequence")?;

    worksheet8.write_dynamic_formula(0, 0, "=SEQUENCE(4,5)")?;

    // -----------------------------------------------------------------------
    // Example of using the Spill range operator.
    // -----------------------------------------------------------------------
    let worksheet9 = workbook.add_worksheet().set_name("Spill ranges")?;

    worksheet9.write_dynamic_formula(1, 7, "=ANCHORARRAY(F2)")?;

    worksheet9.write_dynamic_formula(1, 9, "=COUNTA(ANCHORARRAY(F2))")?;

    // Write the data the to work on.
    worksheet9.write_dynamic_formula(1, 5, "=UNIQUE(B2:B17)")?;
    worksheet9.write_string_with_format(0, 5, "Unique", &header2)?;
    worksheet9.write_string_with_format(0, 7, "Spill", &header2)?;
    worksheet9.write_string_with_format(0, 9, "Spill", &header2)?;

    // Add sample worksheet data to work on.
    write_worksheet_data(worksheet9, &header1)?;
    worksheet9.set_column_width_pixels(4, 20)?;
    worksheet9.set_column_width_pixels(6, 20)?;
    worksheet9.set_column_width_pixels(8, 20)?;

    // -----------------------------------------------------------------------
    // Example of using dynamic ranges with older Excel functions.
    // -----------------------------------------------------------------------
    let worksheet10 = workbook.add_worksheet().set_name("Older functions")?;

    worksheet10.write_dynamic_array_formula(0, 1, 2, 1, "=LEN(A1:A3)")?;

    // Write the data the to work on.
    worksheet10.write_string(0, 0, "Foo")?;
    worksheet10.write_string(1, 0, "Food")?;
    worksheet10.write_string(2, 0, "Frood")?;

    workbook.save("dynamic_arrays.xlsx")?;

    Ok(())
}

// A simple function and data structure to populate some of the worksheets.
fn write_worksheet_data(worksheet: &mut Worksheet, header: &Format) -> Result<(), XlsxError> {
    let worksheet_data = vec![
        ("East", "Tom", "Apple", 6380),
        ("West", "Fred", "Grape", 5619),
        ("North", "Amy", "Pear", 4565),
        ("South", "Sal", "Banana", 5323),
        ("East", "Fritz", "Apple", 4394),
        ("West", "Sravan", "Grape", 7195),
        ("North", "Xi", "Pear", 5231),
        ("South", "Hector", "Banana", 2427),
        ("East", "Tom", "Banana", 4213),
        ("West", "Fred", "Pear", 3239),
        ("North", "Amy", "Grape", 6520),
        ("South", "Sal", "Apple", 1310),
        ("East", "Fritz", "Banana", 6274),
        ("West", "Sravan", "Pear", 4894),
        ("North", "Xi", "Grape", 7580),
        ("South", "Hector", "Apple", 9814),
    ];

    worksheet.write_string_with_format(0, 0, "Region", header)?;
    worksheet.write_string_with_format(0, 1, "Sales Rep", header)?;
    worksheet.write_string_with_format(0, 2, "Product", header)?;
    worksheet.write_string_with_format(0, 3, "Units", header)?;

    let mut row = 1;
    for data in worksheet_data.iter() {
        worksheet.write_string(row, 0, data.0)?;
        worksheet.write_string(row, 1, data.1)?;
        worksheet.write_string(row, 2, data.2)?;
        worksheet.write_number(row, 3, data.3)?;
        row += 1;
    }

    Ok(())
}
