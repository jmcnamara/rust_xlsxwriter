/*!

A cookbook of example programs using `rust_xlsxwriter`.


The examples below can be run as follows:

```bash
git clone git@github.com:jmcnamara/rust_xlsxwriter.git
cd rust_xlsxwriter/
cargo run --example app_demo
```

# Contents

1. [Hello World: Simple getting started example](#hello-world-simple-getting-started-example)
2. [Feature demo: Demonstrates more features of the library](#feature-demo-demonstrates-more-features-of-the-library)
3. [Cell formatting: Demonstrates various formatting options](#cell-formatting-demonstrates-various-formatting-options)
4. [Format colors: Create a palette of the available colors](#format-colors-create-a-palette-of-the-available-colors)
5. [Merging cells: An example of merging cell ranges](#merging-cells-an-example-of-merging-cell-ranges)
6. [Autofilters: Add an autofilter to a worksheet](#autofilters-add-an-autofilter-to-a-worksheet)
7. [Adding worksheet tables](#adding-worksheet-tables)
8. [Rich strings: Add multi-font rich strings to a worksheet](#rich-strings-add-multi-font-rich-strings-to-a-worksheet)
9. [Right to left display: Set a worksheet into right to left display mode](#right-to-left-display-set-a-worksheet-into-right-to-left-display-mode)
10. [Autofitting Columns: Example of autofitting column widths](#autofitting-columns-example-of-autofitting-column-widths)
11. [Insert images: Add images to a worksheet](#insert-images-add-images-to-a-worksheet)
12. [Insert images: Inserting images to fit cell](#insert-images-inserting-images-to-fit-cell)
13. [Adding a watermark: Adding a watermark to a worksheet by adding an image to the header](#adding-a-watermark-adding-a-watermark-to-a-worksheet-by-adding-an-image-to-the-header)
14. [Chart: Simple: Simple getting started chart example](#chart-simple-simple-getting-started-chart-example)
15. [Chart: Area: Excel Area chart example](#chart-area-excel-area-chart-example)
16. [Chart: Bar: Excel Bar chart example](#chart-bar-excel-bar-chart-example)
17. [Chart: Column: Excel Column chart example](#chart-column-excel-column-chart-example)
18. [Chart: Line: Excel Line chart example](#chart-line-excel-line-chart-example)
19. [Chart: Scatter: Excel Scatter chart example](#chart-scatter-excel-scatter-chart-example)
20. [Chart: Pie: Excel Pie chart example](#chart-pie-excel-pie-chart-example)
21. [Chart: Doughnut: Excel Doughnut chart example](#chart-doughnut-excel-doughnut-chart-example)
22. [Chart: Radar: Excel Radar chart example](#chart-radar-excel-radar-chart-example)
23. [Chart: Pattern Fill: Example of a chart with Pattern Fill](#chart-pattern-fill-example-of-a-chart-with-pattern-fill)
24. [Chart: Styles: Example of setting default chart styles](#chart-styles-example-of-setting-default-chart-styles)
25. [Extending generic write() to handle user data types](#extending-generic-write-to-handle-user-data-types)
26. [Defined names: using user defined variable names in worksheets](#defined-names-using-user-defined-variable-names-in-worksheets)
27. [Setting document properties Set the metadata properties for a workbook](#setting-document-properties-set-the-metadata-properties-for-a-workbook)
28. [Headers and Footers: Shows how to set headers and footers](#headers-and-footers-shows-how-to-set-headers-and-footers)
29. [Hyperlinks: Add hyperlinks to a worksheet](#hyperlinks-add-hyperlinks-to-a-worksheet)
30. [Freeze Panes: Example of setting freeze panes in worksheets](#freeze-panes-example-of-setting-freeze-panes-in-worksheets)
31. [Dynamic array formulas: Examples of dynamic arrays and formulas](#dynamic-array-formulas-examples-of-dynamic-arrays-and-formulas)
32. [Excel LAMBDA() function: Example of using the Excel 365 LAMBDA() function](#excel-lambda-function-example-of-using-the-excel-365-lambda-function)
33. [Setting cell protection in a worksheet](#setting-cell-protection-in-a-worksheet)


# Hello World: Simple getting started example

Program to create a simple Hello World style Excel spreadsheet using the
`rust_xlsxwriter`library.

**Image of the output file:**


<img src="https://rustxlsxwriter.github.io/images/hello.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_hello_world.rs

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write a string to cell (0, 0) = A1.
    worksheet.write(0, 0, "Hello")?;

    // Write a number to cell (1, 0) = A2.
    worksheet.write(1, 0, 12345)?;

    // Save the file to disk.
    workbook.save("hello.xlsx")?;

    Ok(())
}
```


# Feature demo: Demonstrates more features of the library

A simple getting started example of some of the features of the`rust_xlsxwriter`
library. It shows some examples of writing different data types, including
dates.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/demo.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_demo.rs

use rust_xlsxwriter::*;

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create some formats to use in the worksheet.
    let bold_format = Format::new().set_bold();
    let decimal_format = Format::new().set_num_format("0.000");
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let merge_format = Format::new()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center);

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set the column width for clarity.
    worksheet.set_column_width(0, 22)?;

    // Write a string without formatting.
    worksheet.write(0, 0, "Hello")?;

    // Write a string with the bold format defined above.
    worksheet.write_with_format(1, 0, "World", &bold_format)?;

    // Write some numbers.
    worksheet.write(2, 0, 1)?;
    worksheet.write(3, 0, 2.34)?;

    // Write a number with formatting.
    worksheet.write_with_format(4, 0, 3.00, &decimal_format)?;

    // Write a formula.
    worksheet.write(5, 0, Formula::new("=SIN(PI()/4)"))?;

    // Write a date.
    let date = ExcelDateTime::from_ymd(2023, 1, 25)?;
    worksheet.write_with_format(6, 0, &date, &date_format)?;

    // Write some links.
    worksheet.write(7, 0, Url::new("https://www.rust-lang.org"))?;
    worksheet.write(8, 0, Url::new("https://www.rust-lang.org").set_text("Rust"))?;

    // Write some merged cells.
    worksheet.merge_range(9, 0, 9, 1, "Merged cells", &merge_format)?;

    // Insert an image.
    let image = Image::new("examples/rust_logo.png")?;
    worksheet.insert_image(1, 2, &image)?;

    // Save the file to disk.
    workbook.save("demo.xlsx")?;

    Ok(())
}
```


# Cell formatting: Demonstrates various formatting options

An example of the various cell formatting options that are available in the
`rust_xlsxwriter`library. These are laid out on worksheets that correspond to the
sections of the Excel "Format Cells" dialog.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/formatting.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_formatting.rs

use rust_xlsxwriter::*;

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a general heading format.
    let header_format = Format::new()
        .set_bold()
        .set_align(FormatAlign::Top)
        .set_border(FormatBorder::Thin)
        .set_background_color(Color::RGB(0xC6EFCE));

    // -----------------------------------------------------------------------
    // Create a worksheet that demonstrates number formatting.
    // -----------------------------------------------------------------------
    let worksheet = workbook.add_worksheet().set_name("Number")?;

    // Make the header row and columns taller and wider for clarity.
    worksheet.set_column_width(0, 18)?;
    worksheet.set_column_width(1, 18)?;
    worksheet.set_row_height(0, 25)?;

    // Add some descriptive text and the formatted numbers.
    worksheet.write_string_with_format(0, 0, "Number Categories", &header_format)?;
    worksheet.write_string_with_format(0, 1, "Formatted Numbers", &header_format)?;

    // Write an unformatted number with the default or "General" format.
    worksheet.write_string(1, 0, "General")?;
    worksheet.write_number(1, 1, 1234.567)?;

    // Write a number with a decimal format.
    worksheet.write_string(2, 0, "Number")?;
    let decimal_format = Format::new().set_num_format("0.00");
    worksheet.write_number_with_format(2, 1, 1234.567, &decimal_format)?;

    // Write a number with a currency format.
    worksheet.write_string(3, 0, "Currency")?;
    let currency_format = Format::new().set_num_format("[$¥-ja-JP]#,##0.00");
    worksheet.write_number_with_format(3, 1, 1234.567, &currency_format)?;

    // Write a number with an accountancy format.
    worksheet.write_string(4, 0, "Accountancy")?;
    let accountancy_format = Format::new().set_num_format("_-[$¥-ja-JP]* #,##0.00_-");
    worksheet.write_number_with_format(4, 1, 1234.567, &accountancy_format)?;

    // Write a number with a short date format.
    worksheet.write_string(5, 0, "Date")?;
    let short_date_format = Format::new().set_num_format("yyyy-mm-dd;@");
    worksheet.write_number_with_format(5, 1, 44927.23, &short_date_format)?;

    // Write a number with a long date format.
    worksheet.write_string(6, 0, "Date")?;
    let long_date_format = Format::new().set_num_format("[$-x-sysdate]dddd, mmmm dd, yyyy");
    worksheet.write_number_with_format(6, 1, 44927.23, &long_date_format)?;

    // Write a number with a percentage format.
    worksheet.write_string(7, 0, "Percentage")?;
    let percentage_format = Format::new().set_num_format("0.00%");
    worksheet.write_number_with_format(7, 1, 72.5 / 100.0, &percentage_format)?;

    // Write a number with a fraction format.
    worksheet.write_string(8, 0, "Fraction")?;
    let fraction_format = Format::new().set_num_format("# ??/??");
    worksheet.write_number_with_format(8, 1, 5.0 / 16.0, &fraction_format)?;

    // Write a number with a percentage format.
    worksheet.write_string(9, 0, "Scientific")?;
    let scientific_format = Format::new().set_num_format("0.00E+00");
    worksheet.write_number_with_format(9, 1, 1234.567, &scientific_format)?;

    // Write a number with a text format.
    worksheet.write_string(10, 0, "Text")?;
    let text_format = Format::new().set_num_format("@");
    worksheet.write_number_with_format(10, 1, 1234.567, &text_format)?;

    // -----------------------------------------------------------------------
    // Create a worksheet that demonstrates number formatting.
    // -----------------------------------------------------------------------
    let worksheet = workbook.add_worksheet().set_name("Alignment")?;

    // Make some rows and columns taller and wider for clarity.
    worksheet.set_column_width(0, 18)?;
    for row_num in 0..5 {
        worksheet.set_row_height(row_num, 30)?;
    }

    // Add some descriptive text at the top of the worksheet.
    worksheet.write_string_with_format(0, 0, "Alignment formats", &header_format)?;

    // Some examples of positional alignment formats.
    let center_format = Format::new().set_align(FormatAlign::Center);
    worksheet.write_string_with_format(1, 0, "Center", &center_format)?;

    let top_left_format = Format::new()
        .set_align(FormatAlign::Top)
        .set_align(FormatAlign::Left);
    worksheet.write_string_with_format(2, 0, "Top - Left", &top_left_format)?;

    let center_center_format = Format::new()
        .set_align(FormatAlign::VerticalCenter)
        .set_align(FormatAlign::Center);
    worksheet.write_string_with_format(3, 0, "Center - Center", &center_center_format)?;

    let bottom_right_format = Format::new()
        .set_align(FormatAlign::Bottom)
        .set_align(FormatAlign::Right);
    worksheet.write_string_with_format(4, 0, "Bottom - Right", &bottom_right_format)?;

    // Some indentation formats.
    let indent1_format = Format::new().set_indent(1);
    worksheet.write_string_with_format(5, 0, "Indent 1", &indent1_format)?;

    let indent2_format = Format::new().set_indent(2);
    worksheet.write_string_with_format(6, 0, "Indent 2", &indent2_format)?;

    // Text wrap format.
    let text_wrap_format = Format::new().set_text_wrap();
    worksheet.write_string_with_format(7, 0, "Some text that is wrapped", &text_wrap_format)?;
    worksheet.write_string_with_format(8, 0, "Text\nwrapped\nat newlines", &text_wrap_format)?;

    // Shrink text format.
    let shrink_format = Format::new().set_shrink();
    worksheet.write_string_with_format(9, 0, "Shrink wide text to fit cell", &shrink_format)?;

    // Text rotation formats.
    let rotate_format1 = Format::new().set_rotation(30);
    worksheet.write_string_with_format(10, 0, "Rotate", &rotate_format1)?;

    let rotate_format2 = Format::new().set_rotation(-30);
    worksheet.write_string_with_format(11, 0, "Rotate", &rotate_format2)?;

    let rotate_format3 = Format::new().set_rotation(270);
    worksheet.write_string_with_format(12, 0, "Rotate", &rotate_format3)?;

    // -----------------------------------------------------------------------
    // Create a worksheet that demonstrates font formatting.
    // -----------------------------------------------------------------------
    let worksheet = workbook.add_worksheet().set_name("Font")?;

    // Make the header row and columns taller and wider for clarity.
    worksheet.set_column_width(0, 18)?;
    worksheet.set_column_width(1, 18)?;
    worksheet.set_row_height(0, 25)?;

    // Add some descriptive text at the top of the worksheet.
    worksheet.write_string_with_format(0, 0, "Font formatting", &header_format)?;

    // Different fonts.
    worksheet.write_string(1, 0, "Calibri 11 (default font)")?;

    let algerian_format = Format::new().set_font_name("Algerian");
    worksheet.write_string_with_format(2, 0, "Algerian", &algerian_format)?;

    let consolas_format = Format::new().set_font_name("Consolas");
    worksheet.write_string_with_format(3, 0, "Consolas", &consolas_format)?;

    let comic_sans_format = Format::new().set_font_name("Comic Sans MS");
    worksheet.write_string_with_format(4, 0, "Comic Sans MS", &comic_sans_format)?;

    // Font styles.
    let bold = Format::new().set_bold();
    worksheet.write_string_with_format(5, 0, "Bold", &bold)?;

    let italic = Format::new().set_italic();
    worksheet.write_string_with_format(6, 0, "Italic", &italic)?;

    let bold_italic = Format::new().set_bold().set_italic();
    worksheet.write_string_with_format(7, 0, "Bold/Italic", &bold_italic)?;

    // Font size.
    let size_format = Format::new().set_font_size(18);
    worksheet.write_string_with_format(8, 0, "Font size 18", &size_format)?;

    // Font color.
    let font_color_format = Format::new().set_font_color(Color::Red);
    worksheet.write_string_with_format(9, 0, "Font color", &font_color_format)?;

    // Font underline.
    let underline_format = Format::new().set_underline(FormatUnderline::Single);
    worksheet.write_string_with_format(10, 0, "Underline", &underline_format)?;

    // Font strike-though.
    let strikethrough_format = Format::new().set_font_strikethrough();
    worksheet.write_string_with_format(11, 0, "Strikethrough", &strikethrough_format)?;

    // -----------------------------------------------------------------------
    // Create a worksheet that demonstrates border formatting.
    // -----------------------------------------------------------------------
    let worksheet = workbook.add_worksheet().set_name("Border")?;

    // Make the header row and columns taller and wider for clarity.
    worksheet.set_column_width(2, 18)?;
    worksheet.set_row_height(0, 25)?;

    // Add some descriptive text at the top of the worksheet.
    worksheet.write_string_with_format(0, 2, "Border formats", &header_format)?;

    // Add some borders to cells.
    let border_format1 = Format::new().set_border(FormatBorder::Thin);
    worksheet.write_string_with_format(2, 2, "Thin Border", &border_format1)?;

    let border_format2 = Format::new().set_border(FormatBorder::Dotted);
    worksheet.write_string_with_format(4, 2, "Dotted Border", &border_format2)?;

    let border_format3 = Format::new().set_border(FormatBorder::Double);
    worksheet.write_string_with_format(6, 2, "Double Border", &border_format3)?;

    let border_format4 = Format::new()
        .set_border(FormatBorder::Thin)
        .set_border_color(Color::Red);
    worksheet.write_string_with_format(8, 2, "Color Border", &border_format4)?;

    // -----------------------------------------------------------------------
    // Create a worksheet that demonstrates fill/pattern formatting.
    // -----------------------------------------------------------------------
    let worksheet = workbook.add_worksheet().set_name("Fill")?;

    // Make the header row and columns taller and wider for clarity.
    worksheet.set_column_width(1, 18)?;
    worksheet.set_row_height(0, 25)?;

    // Add some descriptive text at the top of the worksheet.
    worksheet.write_string_with_format(0, 1, "Fill formats", &header_format)?;

    // Write some cells with pattern fills.
    let fill_format1 = Format::new()
        .set_background_color(Color::Yellow)
        .set_pattern(FormatPattern::Solid);
    worksheet.write_string_with_format(2, 1, "Solid fill", &fill_format1)?;

    let fill_format2 = Format::new()
        .set_background_color(Color::Yellow)
        .set_foreground_color(Color::Orange)
        .set_pattern(FormatPattern::Gray0625);
    worksheet.write_string_with_format(4, 1, "Pattern fill", &fill_format2)?;

    // Save the file to disk.
    workbook.save("cell_formats.xlsx")?;

    Ok(())
}
```


# Format colors: Create a palette of the available colors

This example create a sample palette of the the defined colors, some user
defined RGB colors, and the theme palette color available in the`rust_xlsxwriter`
library.


**Images of the output file:**

<img src="https://rustxlsxwriter.github.io/images/colors.png">

<img src="https://rustxlsxwriter.github.io/images/colors_theme.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_colors.rs

use rust_xlsxwriter::*;

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet for the RGB colors.
    let worksheet = workbook.add_worksheet().set_name("RGB Colors")?;

    // Write some enum defined colors to cells.
    let color_format = Format::new().set_background_color(Color::Black);
    worksheet.write_string(0, 0, "Black")?;
    worksheet.write_blank(0, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Blue);
    worksheet.write_string(1, 0, "Blue")?;
    worksheet.write_blank(1, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Brown);
    worksheet.write_string(2, 0, "Brown")?;
    worksheet.write_blank(2, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Cyan);
    worksheet.write_string(3, 0, "Cyan")?;
    worksheet.write_blank(3, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Gray);
    worksheet.write_string(4, 0, "Gray")?;
    worksheet.write_blank(4, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Green);
    worksheet.write_string(5, 0, "Green")?;
    worksheet.write_blank(5, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Lime);
    worksheet.write_string(6, 0, "Lime")?;
    worksheet.write_blank(6, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Magenta);
    worksheet.write_string(7, 0, "Magenta")?;
    worksheet.write_blank(7, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Navy);
    worksheet.write_string(8, 0, "Navy")?;
    worksheet.write_blank(8, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Orange);
    worksheet.write_string(9, 0, "Orange")?;
    worksheet.write_blank(9, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Pink);
    worksheet.write_string(10, 0, "Pink")?;
    worksheet.write_blank(10, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Purple);
    worksheet.write_string(11, 0, "Purple")?;
    worksheet.write_blank(11, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Red);
    worksheet.write_string(12, 0, "Red")?;
    worksheet.write_blank(12, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Silver);
    worksheet.write_string(13, 0, "Silver")?;
    worksheet.write_blank(13, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::White);
    worksheet.write_string(14, 0, "White")?;
    worksheet.write_blank(14, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::Yellow);
    worksheet.write_string(15, 0, "Yellow")?;
    worksheet.write_blank(15, 1, &color_format)?;

    // Write some user defined RGB colors to cells.
    let color_format = Format::new().set_background_color(Color::RGB(0xFF7F50));
    worksheet.write_string(16, 0, "#FF7F50")?;
    worksheet.write_blank(16, 1, &color_format)?;

    let color_format = Format::new().set_background_color(Color::RGB(0xDCDCDC));
    worksheet.write_string(17, 0, "#DCDCDC")?;
    worksheet.write_blank(17, 1, &color_format)?;

    // Write a RGB color with the shorter Html string variant.
    let color_format = Format::new().set_background_color("#6495ED");
    worksheet.write_string(18, 0, "#6495ED")?;
    worksheet.write_blank(18, 1, &color_format)?;

    // Write a RGB color with the optional u32 variant.
    let color_format = Format::new().set_background_color(0xDAA520);
    worksheet.write_string(19, 0, "#DAA520")?;
    worksheet.write_blank(19, 1, &color_format)?;

    // Add a worksheet for the Theme colors.
    let worksheet = workbook.add_worksheet().set_name("Theme Colors")?;

    // Create a cell with each of the theme colors.
    for row in 0..=5u32 {
        for col in 0..=9u16 {
            let color = col as u8;
            let shade = row as u8;
            let theme_color = Color::Theme(color, shade);
            let text = format!("({}, {})", col, row);

            let mut font_color = Color::White;
            if col == 0 {
                font_color = Color::Default;
            }

            let color_format = Format::new()
                .set_background_color(theme_color)
                .set_font_color(font_color)
                .set_align(FormatAlign::Center);

            worksheet.write_string_with_format(row, col, &text, &color_format)?;
        }
    }

    // Save the file to disk.
    workbook.save("colors.xlsx")?;

    Ok(())
}
```


# Merging cells: An example of merging cell ranges

This is an example of creating merged cells ranges in Excel using
[`worksheet.merge_range()`].

[`worksheet.merge_range()`]: crate::Worksheet::merge_range

The `merge_range()` method only handles strings but it can be used to merge
other data types, such as number, as shown below.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_merge_range.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_merge_range.rs

use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write some merged cells with centering.
    let format = Format::new().set_align(FormatAlign::Center);

    worksheet.merge_range(1, 1, 1, 2, "Merged cells", &format)?;

    // Write some merged cells with centering and a border.
    let format = Format::new()
        .set_align(FormatAlign::Center)
        .set_border(FormatBorder::Thin);

    worksheet.merge_range(3, 1, 3, 2, "Merged cells", &format)?;

    // Write some merged cells with a number by overwriting the first cell in
    // the string merge range with the formatted number.
    worksheet.merge_range(5, 1, 5, 2, "", &format)?;
    worksheet.write_number_with_format(5, 1, 12345.67, &format)?;

    // Example with a more complex format and larger range.
    let format = Format::new()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin)
        .set_background_color(Color::Silver);

    worksheet.merge_range(7, 1, 8, 3, "Merged cells", &format)?;

    // Save the file to disk.
    workbook.save("merge_range.xlsx")?;

    Ok(())
}
```




# Autofilters: Add an autofilter to a worksheet

An example of how to create autofilters with the `rust_xlsxwriter`library..

An autofilter is a way of adding drop down lists to the headers of a 2D range of
worksheet data. This allows users to filter the data based on simple criteria so
that some data is shown and some is hidden.


**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_autofilter1.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_autofilter.rs

use rust_xlsxwriter::{FilterCondition, FilterCriteria, Format, Workbook, Worksheet, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // -----------------------------------------------------------------------
    // 1. Add an autofilter to a data range.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, false)?;

    // Set the autofilter area, including the header/filter row.
    worksheet.autofilter(0, 0, 50, 3)?;

    // -----------------------------------------------------------------------
    // 2. Add an autofilter with a list filter condition.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, false)?;

    // Set the autofilter area.
    worksheet.autofilter(0, 0, 50, 3)?;

    // Set a filter condition to only show cells matching "East" in the first
    // column.
    let filter_condition = FilterCondition::new().add_list_filter("East");
    worksheet.filter_column(0, &filter_condition)?;

    // -----------------------------------------------------------------------
    // 3. Add an autofilter with a list filter condition on multiple items.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, false)?;

    // Set the autofilter area.
    worksheet.autofilter(0, 0, 50, 3)?;

    // Set a filter condition to only show cells matching "East", "West" or
    // "South" in the first column.
    let filter_condition = FilterCondition::new()
        .add_list_filter("East")
        .add_list_filter("West")
        .add_list_filter("South");
    worksheet.filter_column(0, &filter_condition)?;

    // -----------------------------------------------------------------------
    // 4. Add an autofilter with a list filter condition to match blank cells.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, true)?;

    // Set the autofilter area.
    worksheet.autofilter(0, 0, 50, 3)?;

    // Set a filter condition to only show cells matching blanks.
    let filter_condition = FilterCondition::new().add_list_blanks_filter();
    worksheet.filter_column(0, &filter_condition)?;

    // -----------------------------------------------------------------------
    // 5. Add an autofilter with list filters in multiple columns.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, false)?;

    // Set the autofilter area.
    worksheet.autofilter(0, 0, 50, 3)?;

    // Set a filter condition for 2 separate columns.
    let filter_condition1 = FilterCondition::new().add_list_filter("East");
    worksheet.filter_column(0, &filter_condition1)?;

    let filter_condition2 = FilterCondition::new().add_list_filter("July");
    worksheet.filter_column(3, &filter_condition2)?;

    // -----------------------------------------------------------------------
    // 6. Add an autofilter with custom filter condition.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, false)?;

    // Set the autofilter area for numbers greater than 8000.
    worksheet.autofilter(0, 0, 50, 3)?;

    // Set a custom number filter.
    let filter_condition =
        FilterCondition::new().add_custom_filter(FilterCriteria::GreaterThan, 8000);
    worksheet.filter_column(2, &filter_condition)?;

    // -----------------------------------------------------------------------
    // 7. Add an autofilter with 2 custom filters to create a "between" condition.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, false)?;

    // Set the autofilter area.
    worksheet.autofilter(0, 0, 50, 3)?;

    // Set two custom number filters in a "between" configuration.
    let filter_condition = FilterCondition::new()
        .add_custom_filter(FilterCriteria::GreaterThanOrEqualTo, 4000)
        .add_custom_filter(FilterCriteria::LessThanOrEqualTo, 6000);
    worksheet.filter_column(2, &filter_condition)?;

    // -----------------------------------------------------------------------
    // 8. Add an autofilter for non blanks.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, true)?;

    // Set the autofilter area.
    worksheet.autofilter(0, 0, 50, 3)?;

    // Filter non-blanks by filtering on all the unique non-blank
    // strings/numbers in the column.
    let filter_condition = FilterCondition::new()
        .add_list_filter("East")
        .add_list_filter("West")
        .add_list_filter("North")
        .add_list_filter("South");
    worksheet.filter_column(0, &filter_condition)?;

    // Or you can add a simpler custom filter to get the same result.

    // Set a custom number filter of `!= " "` to filter non blanks.
    let filter_condition =
        FilterCondition::new().add_custom_filter(FilterCriteria::NotEqualTo, " ");
    worksheet.filter_column(0, &filter_condition)?;

    // Save the file to disk.
    workbook.save("autofilter.xlsx")?;

    Ok(())
}

// Generate worksheet data to filter on.
pub fn populate_autofilter_data(
    worksheet: &mut Worksheet,
    add_blanks: bool,
) -> Result<(), XlsxError> {
    // The sample data to add to the worksheet.
    let mut data = vec![
        ("East", "Apple", 9000, "July"),
        ("East", "Apple", 5000, "April"),
        ("South", "Orange", 9000, "September"),
        ("North", "Apple", 2000, "November"),
        ("West", "Apple", 9000, "November"),
        ("South", "Pear", 7000, "October"),
        ("North", "Pear", 9000, "August"),
        ("West", "Orange", 1000, "December"),
        ("West", "Grape", 1000, "November"),
        ("South", "Pear", 10000, "April"),
        ("West", "Grape", 6000, "January"),
        ("South", "Orange", 3000, "May"),
        ("North", "Apple", 3000, "December"),
        ("South", "Apple", 7000, "February"),
        ("West", "Grape", 1000, "December"),
        ("East", "Grape", 8000, "February"),
        ("South", "Grape", 10000, "June"),
        ("West", "Pear", 7000, "December"),
        ("South", "Apple", 2000, "October"),
        ("East", "Grape", 7000, "December"),
        ("North", "Grape", 6000, "July"),
        ("East", "Pear", 8000, "February"),
        ("North", "Apple", 7000, "August"),
        ("North", "Orange", 7000, "July"),
        ("North", "Apple", 6000, "June"),
        ("South", "Grape", 8000, "September"),
        ("West", "Apple", 3000, "October"),
        ("South", "Orange", 10000, "November"),
        ("West", "Grape", 4000, "December"),
        ("North", "Orange", 5000, "August"),
        ("East", "Orange", 1000, "November"),
        ("East", "Orange", 4000, "October"),
        ("North", "Grape", 5000, "August"),
        ("East", "Apple", 1000, "July"),
        ("South", "Apple", 10000, "March"),
        ("East", "Grape", 7000, "October"),
        ("West", "Grape", 1000, "September"),
        ("East", "Grape", 10000, "October"),
        ("South", "Orange", 8000, "March"),
        ("North", "Apple", 4000, "July"),
        ("South", "Orange", 5000, "July"),
        ("West", "Apple", 4000, "June"),
        ("East", "Apple", 5000, "April"),
        ("North", "Pear", 3000, "August"),
        ("East", "Grape", 9000, "November"),
        ("North", "Orange", 8000, "October"),
        ("East", "Apple", 10000, "June"),
        ("South", "Pear", 1000, "December"),
        ("North", "Grape", 10000, "July"),
        ("East", "Grape", 6000, "February"),
    ];

    // Introduce blanks cells for some of the examples.
    if add_blanks {
        data[5].0 = "";
        data[18].0 = "";
        data[30].0 = "";
        data[40].0 = "";
    }

    // Widen the columns for clarity.
    worksheet.set_column_width(0, 12)?;
    worksheet.set_column_width(1, 12)?;
    worksheet.set_column_width(2, 12)?;
    worksheet.set_column_width(3, 12)?;

    // Write the header titles.
    let header_format = Format::new().set_bold();
    worksheet.write_string_with_format(0, 0, "Region", &header_format)?;
    worksheet.write_string_with_format(0, 1, "Item", &header_format)?;
    worksheet.write_string_with_format(0, 2, "Volume", &header_format)?;
    worksheet.write_string_with_format(0, 3, "Month", &header_format)?;

    // Write the other worksheet data.
    for (row, data) in data.iter().enumerate() {
        let row = 1 + row as u32;
        worksheet.write_string(row, 0, data.0)?;
        worksheet.write_string(row, 1, data.1)?;
        worksheet.write_number(row, 2, data.2)?;
        worksheet.write_string(row, 3, data.3)?;
    }

    Ok(())
}
```


# Adding worksheet tables

Tables in Excel are a way of grouping a range of cells into a single entity
that has common formatting or that can be referenced from formulas. Tables
can have column headers, autofilters, total rows, column formulas and
different formatting styles.

The image below shows a default table in Excel with the default properties
shown in the ribbon bar.

<img src="https://rustxlsxwriter.github.io/images/table_intro.png">

A table is added to a worksheet via the [`worksheet.add_table()`]method. The
headers and total row of a table should be configured via a [`Table`] struct but
the table data can be added via standard [`worksheet.write()`]methods.

[`Table`]: crate::Table
[`worksheet.write()`]: crate::Worksheet::write
[`worksheet.add_table()`]: crate::Worksheet::add_table

## Some examples:

Example 1. Default table with no data.

<img src="https://rustxlsxwriter.github.io/images/app_tables1.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Default table with no data.";

    // Set the columns widths for clarity.
    for col_num in 1..=6u16 {
        worksheet.set_column_width(col_num, 12)?;
    }

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Create a new table.
    let table = Table::new();

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;
```


Example 2. Default table with data.

<img src="https://rustxlsxwriter.github.io/images/app_tables2.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Default table with data.";

    // Set the columns widths for clarity.
    for col_num in 1..=6u16 {
        worksheet.set_column_width(col_num, 12)?;
    }

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create a new table.
    let table = Table::new();

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;
```


Example 3. Table without default autofilter.

<img src="https://rustxlsxwriter.github.io/images/app_tables3.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table without default autofilter.";

    // Set the columns widths for clarity.
    for col_num in 1..=6u16 {
        worksheet.set_column_width(col_num, 12)?;
    }

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let mut table = Table::new();
    table.set_autofilter(false);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;
```


Example 4. Table without default header row.

<img src="https://rustxlsxwriter.github.io/images/app_tables4.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table without default header row.";

    // Set the columns widths for clarity.
    for col_num in 1..=6u16 {
        worksheet.set_column_width(col_num, 12)?;
    }

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let mut table = Table::new();
    table.set_header_row(false);

    // Add the table to the worksheet.
    worksheet.add_table(3, 1, 6, 5, &table)?;
```


Example 5. Default table with "First Column" and "Last Column" options.

<img src="https://rustxlsxwriter.github.io/images/app_tables5.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Default table with 'First Column' and 'Last Column' options.";

    // Set the columns widths for clarity.
    for col_num in 1..=6u16 {
        worksheet.set_column_width(col_num, 12)?;
    }

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let mut table = Table::new();
    table.set_first_column(true);
    table.set_last_column(true);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;
```


Example 6. Table with banded columns but without default banded rows.

<img src="https://rustxlsxwriter.github.io/images/app_tables6.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with banded columns but without default banded rows.";

    // Set the columns widths for clarity.
    for col_num in 1..=6u16 {
        worksheet.set_column_width(col_num, 12)?;
    }

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let mut table = Table::new();
    table.set_banded_rows(false);
    table.set_banded_columns(true);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;
```


Example 7. Table with user defined column headers.

<img src="https://rustxlsxwriter.github.io/images/app_tables7.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with user defined column headers.";

    // Set the columns widths for clarity.
    for col_num in 1..=6u16 {
        worksheet.set_column_width(col_num, 12)?;
    }

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let mut table = Table::new();
    let columns = vec![
        TableColumn::new().set_header("Product"),
        TableColumn::new().set_header("Quarter 1"),
        TableColumn::new().set_header("Quarter 2"),
        TableColumn::new().set_header("Quarter 3"),
        TableColumn::new().set_header("Quarter 4"),
    ];
    table.set_columns(&columns);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;
```


Example 8. Table with user defined column headers, and formulas.

<img src="https://rustxlsxwriter.github.io/images/app_tables8.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with user defined column headers, and formulas.";

    // Set the columns widths for clarity.
    for col_num in 1..=6u16 {
        worksheet.set_column_width(col_num, 12)?;
    }

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let mut table = Table::new();
    let columns = vec![
        TableColumn::new().set_header("Product"),
        TableColumn::new().set_header("Quarter 1"),
        TableColumn::new().set_header("Quarter 2"),
        TableColumn::new().set_header("Quarter 3"),
        TableColumn::new().set_header("Quarter 4"),
        TableColumn::new()
            .set_header("Year")
            .set_formula("SUM(Table8[@[Quarter 1]:[Quarter 4]])"),
    ];
    table.set_columns(&columns);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 6, &table)?;
```


Example 9. Table with totals row (but no caption or totals).

<img src="https://rustxlsxwriter.github.io/images/app_tables9.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with totals row (but no caption or totals).";

    // Set the columns widths for clarity.
    for col_num in 1..=6u16 {
        worksheet.set_column_width(col_num, 12)?;
    }

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let mut table = Table::new();
    let columns = vec![
        TableColumn::new().set_header("Product"),
        TableColumn::new().set_header("Quarter 1"),
        TableColumn::new().set_header("Quarter 2"),
        TableColumn::new().set_header("Quarter 3"),
        TableColumn::new().set_header("Quarter 4"),
        TableColumn::new()
            .set_header("Year")
            .set_formula("SUM(Table9[@[Quarter 1]:[Quarter 4]])"),
    ];
    table.set_columns(&columns);
    table.set_total_row(true);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 7, 6, &table)?;
```


Example 10. Table with totals row with user captions and functions.

<img src="https://rustxlsxwriter.github.io/images/app_tables10.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with totals row with user captions and functions.";

    // Set the columns widths for clarity.
    for col_num in 1..=6u16 {
        worksheet.set_column_width(col_num, 12)?;
    }

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;
    // Create and configure a new table.
    let mut table = Table::new();
    let columns = vec![
        TableColumn::new()
            .set_header("Product")
            .set_total_label("Totals"),
        TableColumn::new()
            .set_header("Quarter 1")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 2")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 3")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 4")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Year")
            .set_total_function(TableFunction::Sum)
            .set_formula("SUM(Table10[@[Quarter 1]:[Quarter 4]])"),
    ];
    table.set_columns(&columns);
    table.set_total_row(true);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 7, 6, &table)?;

```


Example 11. Table with alternative Excel style.

<img src="https://rustxlsxwriter.github.io/images/app_tables11.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    let worksheet = workbook.add_worksheet();

    let caption = "Table with alternative Excel style.";

    // Set the columns widths for clarity.
    for col_num in 1..=6u16 {
        worksheet.set_column_width(col_num, 12)?;
    }

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let mut table = Table::new();
    let columns = vec![
        TableColumn::new()
            .set_header("Product")
            .set_total_label("Totals"),
        TableColumn::new()
            .set_header("Quarter 1")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 2")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 3")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 4")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Year")
            .set_total_function(TableFunction::Sum)
            .set_formula("SUM(Table11[@[Quarter 1]:[Quarter 4]])"),
    ];
    table.set_columns(&columns);
    table.set_total_row(true);
    table.set_style(TableStyle::Light11);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 7, 6, &table)?;

```


Example 12. Table with Excel style removed.

<img src="https://rustxlsxwriter.github.io/images/app_tables12.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    let worksheet = workbook.add_worksheet();

    let caption = "Table with Excel style removed.";

    // Set the columns widths for clarity.
    for col_num in 1..=6u16 {
        worksheet.set_column_width(col_num, 12)?;
    }

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let mut table = Table::new();
    let columns = vec![
        TableColumn::new()
            .set_header("Product")
            .set_total_label("Totals"),
        TableColumn::new()
            .set_header("Quarter 1")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 2")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 3")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 4")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Year")
            .set_total_function(TableFunction::Sum)
            .set_formula("SUM(Table12[@[Quarter 1]:[Quarter 4]])"),
    ];
    table.set_columns(&columns);
    table.set_total_row(true);
    table.set_style(TableStyle::None);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 7, 6, &table)?;

```


# Rich strings: Add multi-font rich strings to a worksheet

An example of writing "rich" multi-format strings to worksheet cells.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_rich_strings.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_rich_strings.rs

use rust_xlsxwriter::{Color, Format, FormatAlign, FormatScript, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    worksheet.set_column_width(0, 30)?;

    // Add some formats to use in the rich strings.
    let default = Format::default();
    let red = Format::new().set_font_color(Color::Red);
    let blue = Format::new().set_font_color(Color::Blue);
    let bold = Format::new().set_bold();
    let italic = Format::new().set_italic();
    let center = Format::new().set_align(FormatAlign::Center);
    let superscript = Format::new().set_font_script(FormatScript::Superscript);

    // Write some rich strings with multiple formats.
    let segments = [
        (&default, "This is "),
        (&bold, "bold"),
        (&default, " and this is "),
        (&italic, "italic"),
    ];
    worksheet.write_rich_string(0, 0, &segments)?;

    let segments = [
        (&default, "This is "),
        (&red, "red"),
        (&default, " and this is "),
        (&blue, "blue"),
    ];
    worksheet.write_rich_string(2, 0, &segments)?;

    let segments = [
        (&default, "Some "),
        (&bold, "bold text"),
        (&default, " centered"),
    ];
    worksheet.write_rich_string_with_format(4, 0, &segments, &center)?;

    let segments = [(&italic, "j = k"), (&superscript, "(n-1)")];
    worksheet.write_rich_string_with_format(6, 0, &segments, &center)?;

    // It is possible, and idiomatic, to use slices as the string segments.
    let text = "This is blue and this is red";
    let segments = [
        (&default, &text[..8]),
        (&blue, &text[8..12]),
        (&default, &text[12..25]),
        (&red, &text[25..]),
    ];
    worksheet.write_rich_string(8, 0, &segments)?;

    // Save the file to disk.
    workbook.save("rich_strings.xlsx")?;

    Ok(())
}
```


# Right to left display: Set a worksheet into right to left display mode

This is an example of using `rust_xlsxwriter`to create a workbook with the
default worksheet and cell text direction changed from left-to-right to
right-to-left, as required by some middle eastern versions of Excel.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/worksheet_set_right_to_left.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_right_to_left.rs

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add the cell formats.
    let format_left_to_right = Format::new().set_reading_direction(1);
    let format_right_to_left = Format::new().set_reading_direction(2);

    // Add a worksheet in the standard left to right direction.
    let worksheet1 = workbook.add_worksheet();

    // Make the column wider for clarity.
    worksheet1.set_column_width(0, 25)?;

    // Standard direction:         | A1 | B1 | C1 | ...
    worksheet1.write_string(0, 0, "نص عربي / English text")?;
    worksheet1.write_string_with_format(1, 0, "نص عربي / English text", &format_left_to_right)?;
    worksheet1.write_string_with_format(2, 0, "نص عربي / English text", &format_right_to_left)?;

    // Add a worksheet and change it to right to left direction.
    let worksheet2 = workbook.add_worksheet();
    worksheet2.set_right_to_left(true);

    // Make the column wider for clarity.
    worksheet2.set_column_width(0, 25)?;

    // Right to left direction:    ... | C1 | B1 | A1 |
    worksheet2.write_string(0, 0, "نص عربي / English text")?;
    worksheet2.write_string_with_format(1, 0, "نص عربي / English text", &format_left_to_right)?;
    worksheet2.write_string_with_format(2, 0, "نص عربي / English text", &format_right_to_left)?;

    workbook.save("right_to_left.xlsx")?;

    Ok(())
}
```


# Autofitting Columns: Example of autofitting column widths

This is an example of using the simulated autofit option to automatically set
worksheet column widths based on the data in the column.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/autofit.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_autofit.rs

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some worksheet data to demonstrate autofitting.
    worksheet.write_string(0, 0, "Foo")?;
    worksheet.write_string(1, 0, "Food")?;
    worksheet.write_string(2, 0, "Foody")?;
    worksheet.write_string(3, 0, "Froody")?;

    worksheet.write_number(0, 1, 12345)?;
    worksheet.write_number(1, 1, 12345678)?;
    worksheet.write_number(2, 1, 12345)?;

    worksheet.write_string(0, 2, "Some longer text")?;

    worksheet.write_url(0, 3, "http://ww.google.com")?;
    worksheet.write_url(1, 3, "https://github.com")?;

    // Autofit the worksheet.
    worksheet.autofit();

    // Save the file to disk.
    workbook.save("autofit.xlsx")?;

    Ok(())
}
```


# Insert images: Add images to a worksheet

This is an example of a program to insert images into a worksheet.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_images.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_images.rs

use rust_xlsxwriter::{Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Widen the first column to make the text clearer.
    worksheet.set_column_width(0, 30)?;

    // Create a new image object.
    let mut image = Image::new("examples/rust_logo.png")?;

    // Insert the image.
    worksheet.write_string(0, 0, "Insert an image in a cell:")?;
    worksheet.insert_image(0, 1, &image)?;

    // Insert an image offset in the cell.
    worksheet.write_string(7, 0, "Insert an image with an offset:")?;
    worksheet.insert_image_with_offset(7, 1, &image, 5, 5)?;

    // Insert an image with scaling.
    worksheet.write_string(15, 0, "Insert a scaled image:")?;
    image.set_scale_width(0.75).set_scale_height(0.75);
    worksheet.insert_image(15, 1, &image)?;

    // Save the file to disk.
    workbook.save("images.xlsx")?;

    Ok(())
}
```


# Insert images: Inserting images to fit cell

An example of inserting images into a worksheet using `rust_xlsxwriter`so that
they are scaled to a cell. This approach can be useful if you are building up a
spreadsheet of products with a column of images for each product.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_images_fit_to_cell.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_images_fit_to_cell.rs

use rust_xlsxwriter::{Format, FormatAlign, Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    let center = Format::new().set_align(FormatAlign::VerticalCenter);

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Widen the first column to make the text clearer.
    worksheet.set_column_width(0, 30)?;

    // Set larger cells to accommodate the images.
    worksheet.set_column_width_pixels(1, 200)?;
    worksheet.set_row_height_pixels(0, 140)?;
    worksheet.set_row_height_pixels(2, 140)?;
    worksheet.set_row_height_pixels(4, 140)?;

    // Create a new image object.
    let image = Image::new("examples/rust_logo.png")?;

    // Insert the image as standard, without scaling.
    worksheet.write_with_format(0, 0, "Unscaled image inserted into cell:", &center)?;
    worksheet.insert_image(0, 1, &image)?;

    // Insert the image and scale it to fit the entire cell.
    worksheet.write_with_format(2, 0, "Image scaled to fit cell:", &center)?;
    worksheet.insert_image_fit_to_cell(2, 1, &image, false)?;

    // Insert the image and scale it to the cell while maintaining the aspect ratio.
    // In this case it is scaled to the smaller of the width or height scales.
    worksheet.write_with_format(4, 0, "Image scaled with a fixed aspect ratio:", &center)?;
    worksheet.insert_image_fit_to_cell(4, 1, &image, true)?;

    // Save the file to disk.
    workbook.save("images_fit_to_cell.xlsx")?;

    Ok(())
}
```


# Adding a watermark: Adding a watermark to a worksheet by adding an image to the header


An example of adding a worksheet watermark image. This is based on the method of
putting an image in the worksheet header as suggested in the [Microsoft
documentation].

[Microsoft documentation]: https://support.microsoft.com/en-us/office/add-a-watermark-in-excel-a372182a-d733-484e-825c-18ddf3edf009

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_watermark.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_watermark.rs

use rust_xlsxwriter::{HeaderImagePosition, Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let image = Image::new("examples/watermark.png")?;

    // Insert the watermark image in the header.
    worksheet.set_header("&C&[Picture]");
    worksheet.set_header_image(&image, HeaderImagePosition::Center)?;

    // Set Page View mode so the watermark is visible.
    worksheet.set_view_page_layout();

    // Save the file to disk.
    workbook.save("watermark.xlsx")?;

    Ok(())
}
```


# Chart: Simple: Simple getting started chart example

Getting started example of creating simple Excel charts.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/chart.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart.rs


use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some test data for the charts.
    let data = [[1, 2, 3, 4, 5], [2, 4, 6, 8, 10], [3, 6, 9, 12, 15]];
    for (col_num, col_data) in data.iter().enumerate() {
        for (row_num, row_data) in col_data.iter().enumerate() {
            worksheet.write(row_num as u32, col_num as u16, *row_data)?;
        }
    }

    // -----------------------------------------------------------------------
    // Create a new chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Column);

    // Add data series using Excel formula syntax to describe the range.
    chart.add_series().set_values("Sheet1!$A$1:$A$5");
    chart.add_series().set_values("Sheet1!$B$1:$B$5");
    chart.add_series().set_values("Sheet1!$C$1:$C$5");

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 4, &chart)?;

    // -----------------------------------------------------------------------
    // Create another chart to plot the same data as a Line chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Line);

    // Add data series to the chart using a tuple syntax to describe the range.
    // This method is better when you need to create the ranges programmatically
    // to match the data range in the worksheet.
    let row_min = 0;
    let row_max = data[0].len() as u32 - 1;
    chart
        .add_series()
        .set_values(("Sheet1", row_min, 0, row_max, 0));
    chart
        .add_series()
        .set_values(("Sheet1", row_min, 1, row_max, 1));
    chart
        .add_series()
        .set_values(("Sheet1", row_min, 2, row_max, 2));

    // Add the chart to the worksheet.
    worksheet.insert_chart(16, 4, &chart)?;

    workbook.save("chart.xlsx")?;

    Ok(())
}
```


# Chart: Area: Excel Area chart example

Example of creating Excel Area charts.


**Image of the output file:**

Chart 1 in the following example is a default area chart:
<img src="https://rustxlsxwriter.github.io/images/chart_area1.png">

Chart 2 is a stacked area chart:
<img src="https://rustxlsxwriter.github.io/images/chart_area2.png">


Chart 3 is a percentage stacked area chart:
<img src="https://rustxlsxwriter.github.io/images/chart_area3.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_area.rs


use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Number", &bold)?;
    worksheet.write_with_format(0, 1, "Batch 1", &bold)?;
    worksheet.write_with_format(0, 2, "Batch 2", &bold)?;

    let data = [
        [2, 3, 4, 5, 6, 7],
        [40, 40, 50, 30, 25, 50],
        [30, 25, 30, 10, 5, 10],
    ];
    for (col_num, col_data) in data.iter().enumerate() {
        for (row_num, row_data) in col_data.iter().enumerate() {
            worksheet.write(row_num as u32 + 1, col_num as u16, *row_data)?;
        }
    }

    // -----------------------------------------------------------------------
    // Create a new area chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Area);

    // Configure the first data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Add another data series to the chart using the alternative tuple syntax
    // to describe the range. This method is better when you need to create the
    // ranges programmatically to match the data range in the worksheet.
    let row_min = 1;
    let row_max = data[0].len() as u32;
    chart
        .add_series()
        .set_categories(("Sheet1", row_min, 0, row_max, 0))
        .set_values(("Sheet1", row_min, 2, row_max, 2))
        .set_name(("Sheet1", 0, 2));

    // Add a chart title and some axis labels.
    chart.title().set_name("Results of sample analysis");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(11);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a stacked chart sub-type.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::AreaStacked);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title and some axis labels.
    chart.title().set_name("Stacked Chart");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(12);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(17, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a percentage stacked chart sub-type.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::AreaPercentStacked);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title and some axis labels.
    chart.title().set_name("Percent Stacked Chart");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(13);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(33, 3, &chart, 25, 10)?;

    workbook.save("chart_area.xlsx")?;

    Ok(())
}
```


# Chart: Bar: Excel Bar chart example

Example of creating Excel Bar charts.


**Image of the output file:**

Chart 1 in the following example is a default bar chart:
<img src="https://rustxlsxwriter.github.io/images/chart_bar1.png">

Chart 2 is a stacked bar chart:
<img src="https://rustxlsxwriter.github.io/images/chart_bar2.png">

Chart 3 is a percentage stacked bar chart:
<img src="https://rustxlsxwriter.github.io/images/chart_bar3.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_bar.rs


use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Number", &bold)?;
    worksheet.write_with_format(0, 1, "Batch 1", &bold)?;
    worksheet.write_with_format(0, 2, "Batch 2", &bold)?;

    let data = [
        [2, 3, 4, 5, 6, 7],
        [10, 40, 50, 20, 10, 50],
        [30, 60, 70, 50, 40, 30],
    ];
    for (col_num, col_data) in data.iter().enumerate() {
        for (row_num, row_data) in col_data.iter().enumerate() {
            worksheet.write(row_num as u32 + 1, col_num as u16, *row_data)?;
        }
    }

    // -----------------------------------------------------------------------
    // Create a new bar chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Bar);

    // Configure the first data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Add another data series to the chart using the alternative tuple syntax
    // to describe the range. This method is better when you need to create the
    // ranges programmatically to match the data range in the worksheet.
    let row_min = 1;
    let row_max = data[0].len() as u32;
    chart
        .add_series()
        .set_categories(("Sheet1", row_min, 0, row_max, 0))
        .set_values(("Sheet1", row_min, 2, row_max, 2))
        .set_name(("Sheet1", 0, 2));

    // Add a chart title and some axis labels.
    chart.title().set_name("Results of sample analysis");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(11);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a stacked chart sub-type.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::BarStacked);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title and some axis labels.
    chart.title().set_name("Stacked Chart");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(12);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(17, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a percentage stacked chart sub-type.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::BarPercentStacked);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title and some axis labels.
    chart.title().set_name("Percent Stacked Chart");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(13);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(33, 3, &chart, 25, 10)?;

    workbook.save("chart_bar.xlsx")?;

    Ok(())
}
```


# Chart: Column: Excel Column chart example

Example of creating Excel Column charts.


**Image of the output file:**

Chart 1 in the following example is a default column chart:
<img src="https://rustxlsxwriter.github.io/images/chart_column1.png">

Chart 2 is a stacked column chart:
<img src="https://rustxlsxwriter.github.io/images/chart_column2.png">

Chart 3 is a percentage stacked column chart:
<img src="https://rustxlsxwriter.github.io/images/chart_column3.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_column.rs


use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Number", &bold)?;
    worksheet.write_with_format(0, 1, "Batch 1", &bold)?;
    worksheet.write_with_format(0, 2, "Batch 2", &bold)?;

    let data = [
        [2, 3, 4, 5, 6, 7],
        [10, 40, 50, 20, 10, 50],
        [30, 60, 70, 50, 40, 30],
    ];
    for (col_num, col_data) in data.iter().enumerate() {
        for (row_num, row_data) in col_data.iter().enumerate() {
            worksheet.write(row_num as u32 + 1, col_num as u16, *row_data)?;
        }
    }

    // -----------------------------------------------------------------------
    // Create a new column chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Column);

    // Configure the first data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Add another data series to the chart using the alternative tuple syntax
    // to describe the range. This method is better when you need to create the
    // ranges programmatically to match the data range in the worksheet.
    let row_min = 1;
    let row_max = data[0].len() as u32;
    chart
        .add_series()
        .set_categories(("Sheet1", row_min, 0, row_max, 0))
        .set_values(("Sheet1", row_min, 2, row_max, 2))
        .set_name(("Sheet1", 0, 2));

    // Add a chart title and some axis labels.
    chart.title().set_name("Results of sample analysis");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(11);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a stacked chart sub-type.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::ColumnStacked);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title and some axis labels.
    chart.title().set_name("Stacked Chart");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(12);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(17, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a percentage stacked chart sub-type.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::ColumnPercentStacked);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title and some axis labels.
    chart.title().set_name("Percent Stacked Chart");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(13);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(33, 3, &chart, 25, 10)?;

    workbook.save("chart_column.xlsx")?;

    Ok(())
}
```


# Chart: Line: Excel Line chart example

Example of creating Excel Line charts.


**Image of the output file:**

Chart 1 in the following example is a default line chart:
<img src="https://rustxlsxwriter.github.io/images/chart_line1.png">

Chart 2 is a stacked line chart:
<img src="https://rustxlsxwriter.github.io/images/chart_line2.png">

Chart 3 is a percentage stacked line chart:
<img src="https://rustxlsxwriter.github.io/images/chart_line3.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_line.rs


use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Number", &bold)?;
    worksheet.write_with_format(0, 1, "Batch 1", &bold)?;
    worksheet.write_with_format(0, 2, "Batch 2", &bold)?;

    let data = [
        [2, 3, 4, 5, 6, 7],
        [10, 40, 50, 20, 10, 50],
        [30, 60, 70, 50, 40, 30],
    ];
    for (col_num, col_data) in data.iter().enumerate() {
        for (row_num, row_data) in col_data.iter().enumerate() {
            worksheet.write(row_num as u32 + 1, col_num as u16, *row_data)?;
        }
    }

    // -----------------------------------------------------------------------
    // Create a new line chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Line);

    // Configure the first data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Add another data series to the chart using the alternative tuple syntax
    // to describe the range. This method is better when you need to create the
    // ranges programmatically to match the data range in the worksheet.
    let row_min = 1;
    let row_max = data[0].len() as u32;
    chart
        .add_series()
        .set_categories(("Sheet1", row_min, 0, row_max, 0))
        .set_values(("Sheet1", row_min, 2, row_max, 2))
        .set_name(("Sheet1", 0, 2));

    // Add a chart title and some axis labels.
    chart.title().set_name("Results of sample analysis");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(10);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a stacked chart sub-type.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::LineStacked);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title and some axis labels.
    chart.title().set_name("Stacked Chart");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(12);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(17, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a percentage stacked chart sub-type.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::LinePercentStacked);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title and some axis labels.
    chart.title().set_name("Percent Stacked Chart");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(13);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(33, 3, &chart, 25, 10)?;

    workbook.save("chart_line.xlsx")?;

    Ok(())
}
```


# Chart: Scatter: Excel Scatter chart example

Example of creating Excel Scatter charts.


**Image of the output file:**

Chart 1 in the following example is a default scatter chart:
<img src="https://rustxlsxwriter.github.io/images/chart_scatter1.png">

Chart 2 is a scatter chart with straight lines and markers:
<img src="https://rustxlsxwriter.github.io/images/chart_scatter2.png">

Chart 3 is a scatter chart with straight lines and no markers:
<img src="https://rustxlsxwriter.github.io/images/chart_scatter3.png">

Chart 4 is a scatter chart with smooth lines and markers:
<img src="https://rustxlsxwriter.github.io/images/chart_scatter4.png">

Chart 5 is a scatter chart with smooth lines and no markers:
<img src="https://rustxlsxwriter.github.io/images/chart_scatter5.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_scatter.rs


use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Number", &bold)?;
    worksheet.write_with_format(0, 1, "Batch 1", &bold)?;
    worksheet.write_with_format(0, 2, "Batch 2", &bold)?;

    let data = [
        [2, 3, 4, 5, 6, 7],
        [10, 40, 50, 20, 10, 50],
        [30, 60, 70, 50, 40, 30],
    ];
    for (col_num, col_data) in data.iter().enumerate() {
        for (row_num, row_data) in col_data.iter().enumerate() {
            worksheet.write(row_num as u32 + 1, col_num as u16, *row_data)?;
        }
    }

    // -----------------------------------------------------------------------
    // Create a new scatter chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Scatter);

    // Configure the first data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Add another data series to the chart using the alternative tuple syntax
    // to describe the range. This method is better when you need to create the
    // ranges programmatically to match the data range in the worksheet.
    let row_min = 1;
    let row_max = data[0].len() as u32;
    chart
        .add_series()
        .set_categories(("Sheet1", row_min, 0, row_max, 0))
        .set_values(("Sheet1", row_min, 2, row_max, 2))
        .set_name(("Sheet1", 0, 2));

    // Add a chart title and some axis labels.
    chart.title().set_name("Results of sample analysis");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(11);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a scatter chart sub-type with straight lines and markers.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::ScatterStraightWithMarkers);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title and some axis labels.
    chart.title().set_name("Straight line with markers");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(12);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(17, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a scatter chart sub-type with straight lines and no markers.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::ScatterStraight);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title and some axis labels.
    chart.title().set_name("Straight line");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(13);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(33, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a scatter chart sub-type with smooth lines and markers.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::ScatterSmoothWithMarkers);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title and some axis labels.
    chart.title().set_name("Smooth line with markers");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(14);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(49, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a scatter chart sub-type with smooth lines and no markers.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::ScatterSmooth);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title and some axis labels.
    chart.title().set_name("Smooth line");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(15);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(65, 3, &chart, 25, 10)?;

    workbook.save("chart_scatter.xlsx")?;

    Ok(())
}
```


# Chart: Pie: Excel Pie chart example

Example of creating Excel Pie charts.


**Image of the output file:**

Chart 1 in the following example is a default pie chart:
<img src="https://rustxlsxwriter.github.io/images/chart_pie1.png">

Chart 2 shows how to set segment colors.

It is possible to define chart colors for most types of `rust_xlsxwriter`charts via
the `add_series()` method. However, Pie charts are a special case since each
segment is represented as a point and as such it is necessary to assign
formatting to each point in the series.
<img src="https://rustxlsxwriter.github.io/images/chart_pie2.png">

Chart 3 shows how to rotate the segments of the chart:
<img src="https://rustxlsxwriter.github.io/images/chart_pie3.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_pie.rs


use rust_xlsxwriter::{Chart, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Category", &bold)?;
    worksheet.write_with_format(0, 1, "Values", &bold)?;

    worksheet.write(1, 0, "Apple")?;
    worksheet.write(2, 0, "Cherry")?;
    worksheet.write(3, 0, "Pecan")?;

    worksheet.write(1, 1, 60)?;
    worksheet.write(2, 1, 30)?;
    worksheet.write(3, 1, 10)?;

    // -----------------------------------------------------------------------
    // Create a new pie chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new_pie();

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$4")
        .set_values("Sheet1!$B$2:$B$4")
        .set_name("Pie sales data");

    // Add a chart title.
    chart.title().set_name("Popular Pie Types");

    // Set an Excel chart style.
    chart.set_style(10);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 2, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a Pie chart with user defined segment colors.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new_pie();

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$4")
        .set_values("Sheet1!$B$2:$B$4")
        .set_name("Pie sales data")
        .set_point_colors(&["#5ABA10", "#FE110E", "#CA5C05"]);

    // Add a chart title.
    chart.title().set_name("Pie Chart with user defined colors");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(17, 2, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a Pie chart with rotation of the segments.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new_pie();

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$4")
        .set_values("Sheet1!$B$2:$B$4")
        .set_name("Pie sales data");

    // Change the angle/rotation of the first segment.
    chart.set_rotation(90);

    // Add a chart title.
    chart.title().set_name("Pie Chart with segment rotation");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(33, 2, &chart, 25, 10)?;

    workbook.save("chart_pie.xlsx")?;

    Ok(())
}
```


# Chart: Doughnut: Excel Doughnut chart example

Example of creating Excel Doughnut charts.


**Image of the output file:**

Chart 1 in the following example is a default doughnut chart:
<img src="https://rustxlsxwriter.github.io/images/chart_doughnut1.png">




Chart 4 shows how to set segment colors and other options.

<img src="https://rustxlsxwriter.github.io/images/chart_doughnut2.png">



**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_doughnut.rs


use rust_xlsxwriter::{Chart, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Category", &bold)?;
    worksheet.write_with_format(0, 1, "Values", &bold)?;

    worksheet.write(1, 0, "Glazed")?;
    worksheet.write(2, 0, "Chocolate")?;
    worksheet.write(3, 0, "Cream")?;

    worksheet.write(1, 1, 50)?;
    worksheet.write(2, 1, 35)?;
    worksheet.write(3, 1, 15)?;

    // -----------------------------------------------------------------------
    // Create a new doughnut chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new_doughnut();

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$4")
        .set_values("Sheet1!$B$2:$B$4")
        .set_name("Doughnut sales data");

    // Add a chart title.
    chart.title().set_name("Popular Doughnut Types");

    // Set an Excel chart style.
    chart.set_style(10);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 2, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a Doughnut chart with user defined segment colors.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new_doughnut();

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$4")
        .set_values("Sheet1!$B$2:$B$4")
        .set_name("Doughnut sales data")
        .set_point_colors(&["#FA58D0", "#61210B", "#F5F6CE"]);

    // Add a chart title.
    chart
        .title()
        .set_name("Doughnut Chart with user defined colors");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(17, 2, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a Doughnut chart with rotation of the segments.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new_doughnut();

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$4")
        .set_values("Sheet1!$B$2:$B$4")
        .set_name("Doughnut sales data");

    // Change the angle/rotation of the first segment.
    chart.set_rotation(90);

    // Add a chart title.
    chart
        .title()
        .set_name("Doughnut Chart with segment rotation");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(33, 2, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a Doughnut chart with user defined hole size and other options.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new_doughnut();

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$4")
        .set_values("Sheet1!$B$2:$B$4")
        .set_name("Doughnut sales data")
        .set_point_colors(&["#FA58D0", "#61210B", "#F5F6CE"]);

    // Add a chart title.
    chart
        .title()
        .set_name("Doughnut Chart with options applied");

    // Change the angle/rotation of the first segment.
    chart.set_rotation(28);

    // Change the hole size.
    chart.set_hole_size(33);

    // Set a 3D style.
    chart.set_style(26);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(49, 2, &chart, 25, 10)?;

    workbook.save("chart_doughnut.xlsx")?;

    Ok(())
}
```


# Chart: Radar: Excel Radar chart example

Example of creating Excel Radar charts.


**Image of the output file:**

Chart 1 in the following example is a default radar chart:
<img src="https://rustxlsxwriter.github.io/images/chart_radar1.png">

Chart 2 in the following example is a radar chart with markers:
<img src="https://rustxlsxwriter.github.io/images/chart_radar2.png">

Chart 3 in the following example is a filled radar chart:
<img src="https://rustxlsxwriter.github.io/images/chart_radar3.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_radar.rs


use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Number", &bold)?;
    worksheet.write_with_format(0, 1, "Batch 1", &bold)?;
    worksheet.write_with_format(0, 2, "Batch 2", &bold)?;

    let data = [
        [2, 3, 4, 5, 6, 7],
        [30, 60, 70, 50, 40, 30],
        [25, 40, 50, 30, 50, 40],
    ];
    for (col_num, col_data) in data.iter().enumerate() {
        for (row_num, row_data) in col_data.iter().enumerate() {
            worksheet.write(row_num as u32 + 1, col_num as u16, *row_data)?;
        }
    }

    // -----------------------------------------------------------------------
    // Create a new radar chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Radar);

    // Configure the first data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Add another data series to the chart using the alternative tuple syntax
    // to describe the range. This method is better when you need to create the
    // ranges programmatically to match the data range in the worksheet.
    let row_min = 1;
    let row_max = data[0].len() as u32;
    chart
        .add_series()
        .set_categories(("Sheet1", row_min, 0, row_max, 0))
        .set_values(("Sheet1", row_min, 2, row_max, 2))
        .set_name(("Sheet1", 0, 2));

    // Add a chart title.
    chart.title().set_name("Results of sample analysis");

    // Set an Excel chart style.
    chart.set_style(11);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a radar chart with markers.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::RadarWithMarkers);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title.
    chart.title().set_name("Radar Chart With Markers");

    // Set an Excel chart style.
    chart.set_style(12);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(17, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a filled radar chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::RadarFilled);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title.
    chart.title().set_name("Filled Radar Chart");

    // Set an Excel chart style.
    chart.set_style(13);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(33, 3, &chart, 25, 10)?;

    workbook.save("chart_radar.xlsx")?;

    Ok(())
}
```


# Chart: Pattern Fill: Example of a chart with Pattern Fill

an example of creating column charts with fill patterns using the [`ChartFormat`]
and [`ChartPatternFill`] structs.


**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/chart_pattern.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_pattern.rs

use rust_xlsxwriter::*;

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Shingle", &bold)?;
    worksheet.write_with_format(0, 1, "Brick", &bold)?;

    let data = [[105, 150, 130, 90], [50, 120, 100, 110]];
    for (col_num, col_data) in data.iter().enumerate() {
        for (row_num, row_data) in col_data.iter().enumerate() {
            worksheet.write(row_num as u32 + 1, col_num as u16, *row_data)?;
        }
    }

    // Create a new column chart.
    let mut chart = Chart::new(ChartType::Column);

    // Configure the first data series and add fill patterns.
    chart
        .add_series()
        .set_name("Sheet1!$A$1")
        .set_values("Sheet1!$A$2:$A$5")
        .set_gap(70)
        .set_format(
            ChartFormat::new()
                .set_pattern_fill(
                    ChartPatternFill::new()
                        .set_pattern(ChartPatternFillType::Shingle)
                        .set_foreground_color("#804000")
                        .set_background_color("#C68C53"),
                )
                .set_border(ChartLine::new().set_color("#804000")),
        );

    chart
        .add_series()
        .set_name("Sheet1!$B$1")
        .set_values("Sheet1!$B$2:$B$5")
        .set_format(
            ChartFormat::new()
                .set_pattern_fill(
                    ChartPatternFill::new()
                        .set_pattern(ChartPatternFillType::HorizontalBrick)
                        .set_foreground_color("#B30000")
                        .set_background_color("#FF6666"),
                )
                .set_border(ChartLine::new().set_color("#B30000")),
        );

    // Add a chart title and some axis labels.
    chart.title().set_name("Cladding types");
    chart.x_axis().set_name("Region");
    chart.y_axis().set_name("Number of houses");

    // Add the chart to the worksheet.
    worksheet.insert_chart(1, 3, &chart)?;

    workbook.save("chart_pattern.xlsx")?;

    Ok(())
}
```

[`ChartFormat`]: crate::ChartFormat
[`ChartPatternFill`]: crate::ChartPatternFill


# Chart: Styles: Example of setting default chart styles

An example showing all 48 default chart styles available in Excel 2007 using the
chart `set_style()` method.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/chart_styles.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_styles.rs

use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let chart_types = vec![
        ("Column", ChartType::Column),
        ("Area", ChartType::Area),
        ("Line", ChartType::Line),
        ("Pie", ChartType::Pie),
    ];

    // Create a worksheet with 48 charts in each of the available styles, for
    // each of the chart types above.
    for (name, chart_type) in chart_types {
        let worksheet = workbook.add_worksheet().set_name(name)?.set_zoom(30);
        let mut chart = Chart::new(chart_type);
        chart.add_series().set_values("Data!$A$1:$A$6");
        let mut style = 1;

        for row_num in (0..90).step_by(15) {
            for col_num in (0..64).step_by(8) {
                chart.set_style(style);
                chart.title().set_name(&format!("Style {style}"));
                chart.legend().set_hidden();
                worksheet.insert_chart(row_num as u32, col_num as u16, &chart)?;
                style += 1;
            }
        }
    }

    // Create a worksheet with data for the charts.
    let data_worksheet = workbook.add_worksheet().set_name("Data")?;
    data_worksheet.write(0, 0, 10)?;
    data_worksheet.write(1, 0, 40)?;
    data_worksheet.write(2, 0, 50)?;
    data_worksheet.write(3, 0, 20)?;
    data_worksheet.write(4, 0, 10)?;
    data_worksheet.write(5, 0, 50)?;
    data_worksheet.set_hidden(true);

    workbook.save("chart_styles.xlsx")?;

    Ok(())
}
```


# Extending generic write() to handle user data types

Example of how to extend the the `rust_xlsxwriter`[`worksheet.write()`] method using the
[`IntoExcelData`] trait to handle arbitrary user data that can be mapped to one
of the main Excel data types.

For this example we create a simple struct type to represent a [Unix Time]. This
is the number of elapsed seconds since the epoch of January 1970 (UTC). Note,
this is for demonstration purposes only. The [`ExcelDateTime`] struct in
 `rust_xlsxwriter` can handle Unix timestamps.


[Unix Time]: https://en.wikipedia.org/wiki/Unix_time
[`IntoExcelData`]: crate::IntoExcelData
[`ExcelDateTime`]: crate::ExcelDateTime
[`worksheet.write()`]: crate::Worksheet::write

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/write_generic.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_write_generic_data.rs

use rust_xlsxwriter::*;

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add a format for the dates.
    let format = Format::new().set_num_format("yyyy-mm-dd");

    // Make the first column wider for clarity.
    worksheet.set_column_width(0, 12)?;

    // Write user defined type instances that implement the IntoExcelData trait.
    worksheet.write_with_format(0, 0, UnixTime::new(0), &format)?;
    worksheet.write_with_format(1, 0, UnixTime::new(946598400), &format)?;
    worksheet.write_with_format(2, 0, UnixTime::new(1672531200), &format)?;

    // Save the file to disk.
    workbook.save("write_generic.xlsx")?;

    Ok(())
}

// For this example we create a simple struct type to represent a Unix time.
// This is the number of elapsed seconds since the epoch of January 1970 (UTC).
// See https://en.wikipedia.org/wiki/Unix_time. Note, this is for demonstration
// purposes only. The `ExcelDateTime` struct in `rust_xlsxwriter` can handle
// Unix timestamps.
pub struct UnixTime {
    seconds: u64,
}

impl UnixTime {
    pub fn new(seconds: u64) -> UnixTime {
        UnixTime { seconds }
    }
}

// Implement the IntoExcelData trait to map our new UnixTime struct into an
// Excel type.
//
// The relevant Excel type is f64 which is used to store dates and times (along
// with a number format). The Unix 1970 epoch equates to a date/number of
// 25569.0. For Unix times beyond that we divide by the number of seconds in the
// day (24 * 60 * 60) to get the Excel serial date.
//
// We need to implement two methods for the trait in order to write data with
// and without a format.
//
impl IntoExcelData for UnixTime {
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Convert the Unix time to an Excel datetime.
        let datetime = 25569.0 + (self.seconds as f64 / (24.0 * 60.0 * 60.0));

        // Write the date as a number with a format.
        worksheet.write_number(row, col, datetime)
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &'a Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        // Convert the Unix time to an Excel datetime.
        let datetime = 25569.0 + (self.seconds as f64 / (24.0 * 60.0 * 60.0));

        // Write the date with the user supplied format.
        worksheet.write_number_with_format(row, col, datetime, format)
    }
}
```


# Defined names: using user defined variable names in worksheets

Example of how to create defined names using the `rust_xlsxwriter` library.

This functionality is used to define user friendly variable names to represent a
value, a single cell,  or a range of cells in a workbook.

**Images of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_defined_name1.png">

Here is the output in the Excel Name Manager. Note that there is a
Global/Workbook "Sales" variable name and a Local/Worksheet version.

<img src="https://rustxlsxwriter.github.io/images/app_defined_name2.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_defined_name.rs

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add two worksheets to the workbook.
    let _worksheet1 = workbook.add_worksheet();
    let _worksheet2 = workbook.add_worksheet();

    // Define some global/workbook names.
    workbook.define_name("Exchange_rate", "=0.96")?;
    workbook.define_name("Sales", "=Sheet1!$G$1:$H$10")?;

    // Define a local/worksheet name. Over-rides the "Sales" name above.
    workbook.define_name("Sheet2!Sales", "=Sheet2!$G$1:$G$10")?;

    // Write some text in the file and one of the defined names in a formula.
    for worksheet in workbook.worksheets_mut() {
        worksheet.set_column_width(0, 45)?;
        worksheet.write_string(0, 0, "This worksheet contains some defined names.")?;
        worksheet.write_string(1, 0, "See Formulas -> Name Manager above.")?;
        worksheet.write_string(2, 0, "Example formula in cell B3 ->")?;

        worksheet.write_formula(2, 1, "=Exchange_rate")?;
    }

    // Save the file to disk.
    workbook.save("defined_name.xlsx")?;

    Ok(())
}
```


# Setting document properties Set the metadata properties for a workbook

An example of setting workbook document properties for a file created using the
`rust_xlsxwriter`library.

**Image of the output file:**


<img src="https://rustxlsxwriter.github.io/images/app_doc_properties.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_doc_properties.rs

use rust_xlsxwriter::{DocProperties, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let properties = DocProperties::new()
        .set_title("This is an example spreadsheet")
        .set_subject("That demonstrates document properties")
        .set_author("A. Rust User")
        .set_manager("J. Alfred Prufrock")
        .set_company("Rust Solutions Inc")
        .set_category("Sample spreadsheets")
        .set_keywords("Sample, Example, Properties")
        .set_comment("Created with Rust and rust_xlsxwriter");

    workbook.set_properties(&properties);

    let worksheet = workbook.add_worksheet();

    worksheet.set_column_width(0, 30)?;
    worksheet.write_string(0, 0, "See File -> Info -> Properties")?;

    workbook.save("doc_properties.xlsx")?;

    Ok(())
}
```


# Headers and Footers: Shows how to set headers and footers


This program shows several examples of how to set up worksheet headers and footers.


Here are some examples from the code and the relevant Excel output.

Some simple text:

```ignore
    worksheet1.set_header("&CHere is some centered text.");
```
<img src="https://rustxlsxwriter.github.io/images/app_header_example1.png">

An example with variables:

```ignore
    worksheet2.set_header("&LPage &[Page] of &[Pages]&CFilename: &[File]&RSheetname: &[Tab]");
```
<img src="https://rustxlsxwriter.github.io/images/app_header_example2.png">


An example of setting a header image:

```ignore
    worksheet3.set_header("&L&[Picture]");
    worksheet3.set_header_image(&image, HeaderImagePosition::Left)?;
```
<img src="https://rustxlsxwriter.github.io/images/app_header_example3.png">


An example of how to use more than one font:

```ignore
    worksheet4.set_header(r#"&C&"Courier New,Bold"Hello &"Arial,Italic"World"#);
```
<img src="https://rustxlsxwriter.github.io/images/app_header_example4.png">

An example of line wrapping:

```ignore
    worksheet5.set_header("&CHeading 1\nHeading 2");
```
<img src="https://rustxlsxwriter.github.io/images/app_header_example5.png">

An example of inserting a literal ampersand &:

```ignore
    worksheet6.set_header("&CCuriouser && Curiouser - Attorneys at Law");
```
<img src="https://rustxlsxwriter.github.io/images/app_header_example6.png">

And here is the full code for the example:

```rust
// Sample code from examples/app_headers_footers.rs

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

    let mut image = Image::new("examples/rust_logo.png")?;
    image.set_scale_height(0.5);
    image.set_scale_width(0.5);

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
```


# Hyperlinks: Add hyperlinks to a worksheet

This is an example of a program to create demonstrate creating links in a
worksheet using the `rust_xlsxwriter`library.

The links can be to external urls, to external files or internally to cells in
the workbook.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_hyperlinks.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_hyperlinks.rs

use rust_xlsxwriter::{Color, Format, FormatUnderline, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create a format to use in the worksheet.
    let link_format = Format::new()
        .set_font_color(Color::Red)
        .set_underline(FormatUnderline::Single);

    // Add a worksheet to the workbook.
    let worksheet1 = workbook.add_worksheet();

    // Set the column width for clarity.
    worksheet1.set_column_width(0, 26)?;

    // Write some url links.
    worksheet1.write_url(0, 0, "https://www.rust-lang.org")?;
    worksheet1.write_url_with_text(1, 0, "https://www.rust-lang.org", "Learn Rust")?;
    worksheet1.write_url_with_format(2, 0, "https://www.rust-lang.org", &link_format)?;

    // Write some internal links.
    worksheet1.write_url(4, 0, "internal:Sheet1!A1")?;
    worksheet1.write_url(5, 0, "internal:Sheet2!C4")?;

    // Write some external links.
    worksheet1.write_url(7, 0, r"file:///C:\Temp\Book1.xlsx")?;
    worksheet1.write_url(8, 0, r"file:///C:\Temp\Book1.xlsx#Sheet1!C4")?;

    // Add another sheet to link to.
    let worksheet2 = workbook.add_worksheet();
    worksheet2.write_string(3, 2, "Here I am")?;
    worksheet2.write_url_with_text(4, 2, "internal:Sheet1!A6", "Go back")?;

    // Save the file to disk.
    workbook.save("hyperlinks.xlsx")?;

    Ok(())
}
```


# Freeze Panes: Example of setting freeze panes in worksheets

An example of setting some "freeze" panes in worksheets to split the worksheet
into scrolling and non-scrolling areas. This is generally used to have one or
more row or column to the top or left of the worksheet area that stays fixed
when a user scrolls.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/panes.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_panes.rs

use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create some formats to use in the worksheet.
    let header_format = Format::new()
        .set_bold()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_foreground_color(Color::RGB(0xD7E4BC))
        .set_border(FormatBorder::Thin);

    let center_format = Format::new().set_align(FormatAlign::Center);

    // Some range limits to use in this example.
    let max_row = 50;
    let max_col = 26;

    // -----------------------------------------------------------------------
    // Example 1. Freeze pane on the top row.
    // -----------------------------------------------------------------------
    let worksheet1 = workbook.add_worksheet().set_name("Panes 1")?;

    // Freeze the top row only.
    worksheet1.set_freeze_panes(1, 0)?;

    // Add some data and formatting to the worksheet.
    worksheet1.set_row_height(0, 20)?;
    for col in 0..max_col {
        worksheet1.write_string_with_format(0, col, "Scroll down", &header_format)?;
        worksheet1.set_column_width(col, 16)?;
    }
    for row in 1..max_row {
        for col in 0..max_col {
            worksheet1.write_number_with_format(row, col, row + 1, &center_format)?;
        }
    }

    // -----------------------------------------------------------------------
    // Example 2. Freeze pane on the left column.
    // -----------------------------------------------------------------------
    let worksheet2 = workbook.add_worksheet().set_name("Panes 2")?;

    // Freeze the leftmost column only.
    worksheet2.set_freeze_panes(0, 1)?;

    // Add some data and formatting to the worksheet.
    worksheet2.set_column_width(0, 16)?;
    for row in 0..max_row {
        worksheet2.write_string_with_format(row, 0, "Scroll Across", &header_format)?;

        for col in 1..max_col {
            worksheet2.write_number_with_format(row, col, col, &center_format)?;
        }
    }

    // -----------------------------------------------------------------------
    // Example 3. Freeze pane on the top row and leftmost column.
    // -----------------------------------------------------------------------
    let worksheet3 = workbook.add_worksheet().set_name("Panes 3")?;

    // Freeze the top row and leftmost column.
    worksheet3.set_freeze_panes(1, 1)?;

    // Add some data and formatting to the worksheet.
    worksheet3.set_row_height(0, 20)?;
    worksheet3.set_column_width(0, 16)?;
    worksheet3.write_blank(0, 0, &header_format)?;

    for col in 1..max_col {
        worksheet3.write_string_with_format(0, col, "Scroll down", &header_format)?;
        worksheet3.set_column_width(col, 16)?;
    }
    for row in 1..max_row {
        worksheet3.write_string_with_format(row, 0, "Scroll Across", &header_format)?;

        for col in 1..max_col {
            worksheet3.write_number_with_format(row, col, col, &center_format)?;
        }
    }

    // -----------------------------------------------------------------------
    // Example 4. Freeze pane on the top row and leftmost column, with
    //            scrolling area shifted.
    // -----------------------------------------------------------------------
    let worksheet4 = workbook.add_worksheet().set_name("Panes 4")?;

    // Freeze the top row and leftmost column.
    worksheet4.set_freeze_panes(1, 1)?;

    // Shift the scrolled area in the scrolling pane.
    worksheet4.set_freeze_panes_top_cell(20, 12)?;

    // Add some data and formatting to the worksheet.
    worksheet4.set_row_height(0, 20)?;
    worksheet4.set_column_width(0, 16)?;
    worksheet4.write_blank(0, 0, &header_format)?;

    for col in 1..max_col {
        worksheet4.write_string_with_format(0, col, "Scroll down", &header_format)?;
        worksheet4.set_column_width(col, 16)?;
    }
    for row in 1..max_row {
        worksheet4.write_string_with_format(row, 0, "Scroll Across", &header_format)?;

        for col in 1..max_col {
            worksheet4.write_number_with_format(row, col, col, &center_format)?;
        }
    }

    // Save the file to disk.
    workbook.save("panes.xlsx")?;

    Ok(())
}
```


# Dynamic array formulas: Examples of dynamic arrays and formulas

An example of how to use the `rust_xlsxwriter`library to write formulas and
functions that create dynamic arrays. These functions are new to Excel
365. The examples mirror the examples in the Excel documentation for these
functions.

**Images of the output file:**


<img src="https://rustxlsxwriter.github.io/images/dynamic_arrays01.png">

Here is another example:

<img src="https://rustxlsxwriter.github.io/images/dynamic_arrays02.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_dynamic_arrays.rs

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
```


# Excel LAMBDA() function: Example of using the Excel 365 LAMBDA() function

An example of using the new Excel `LAMBDA()` function with the`rust_xlsxwriter`
library.


**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_lambda.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_lambda.rs

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Write a Lambda function to convert Fahrenheit to Celsius to a cell as a
    // defined name and use that to calculate a value.
    //
    // Note that the formula name is prefixed with "_xlfn." (this is normally
    // converted automatically by write_formula*() but isn't for defined names)
    // and note that the lambda function parameters are prefixed with "_xlpm.".
    // These prefixes won't show up in Excel.
    workbook.define_name(
        "ToCelsius",
        "=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))",
    )?;

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write the same Lambda function as a cell formula.
    //
    // Note that the lambda function parameters must be prefixed with "_xlpm.".
    // These prefixes won't show up in Excel.
    worksheet.write_formula(0, 0, "=LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))(32)")?;

    // The user defined name needs to be written explicitly as a dynamic array
    // formula.
    worksheet.write_dynamic_formula(1, 0, "=ToCelsius(212)")?;

    // Save the file to disk.
    workbook.save("lambda.xlsx")?;

    Ok(())
}
```


# Setting cell protection in a worksheet

Example of cell locking and formula hiding in an Excel worksheet using worksheet
protection.

**Image of the output file:**


<img src="https://rustxlsxwriter.github.io/images/app_worksheet_protection.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_worksheet_protection.rs

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create some format objects.
    let unlocked = Format::new().set_unlocked();
    let hidden = Format::new().set_hidden();

    // Protect the worksheet to turn on cell locking.
    worksheet.protect();

    // Examples of cell locking and hiding.
    worksheet.write_string(0, 0, "Cell B1 is locked. It cannot be edited.")?;
    worksheet.write_formula(0, 1, "=1+2")?; // Locked by default.

    worksheet.write_string(1, 0, "Cell B2 is unlocked. It can be edited.")?;
    worksheet.write_formula_with_format(1, 1, "=1+2", &unlocked)?;

    worksheet.write_string(2, 0, "Cell B3 is hidden. The formula isn't visible.")?;
    worksheet.write_formula_with_format(2, 1, "=1+2", &hidden)?;

    worksheet.write_string(4, 0, "Use Menu -> Review -> Unprotect Sheet")?;
    worksheet.write_string(5, 0, "to remove the worksheet protection.")?;

    worksheet.autofit();

    // Save the file to disk.
    workbook.save("worksheet_protection.xlsx")?;

    Ok(())
}
```


*/
