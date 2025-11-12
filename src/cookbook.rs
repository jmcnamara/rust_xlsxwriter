/*!

A cookbook of example programs using `rust_xlsxwriter`.


The examples below can be run as follows:

```bash
git clone git@github.com:jmcnamara/rust_xlsxwriter.git
cd rust_xlsxwriter/
cargo run --example app_demo  # or any other example
```

# Contents

1. [Hello World: Simple getting started example](#hello-world-simple-getting-started-example)
2. [Feature demo: Demonstrates more features of the library](#feature-demo-demonstrates-more-features-of-the-library)
3. [Cell formatting: Demonstrates various formatting options](#cell-formatting-demonstrates-various-formatting-options)
4. [Format colors: Create a palette of the available colors](#format-colors-create-a-palette-of-the-available-colors)
5. [Merging cells: An example of merging cell ranges](#merging-cells-an-example-of-merging-cell-ranges)
6. [Autofilters: Add an autofilter to a worksheet](#autofilters-add-an-autofilter-to-a-worksheet)
7. [Tables: Adding worksheet tables](#tables-adding-worksheet-tables)
8. [Conditional Formatting: Adding conditional formatting to worksheets](#conditional-formatting-adding-conditional-formatting-to-worksheets)
9. [Data Validation: Add cell validation and dropdowns](#data-validation-add-cell-validation-and-dropdowns)
10. [Notes: Adding notes to worksheet cells](#notes-adding-notes-to-worksheet-cells)
11. [Rich strings: Add multi-font rich strings to a worksheet](#rich-strings-add-multi-font-rich-strings-to-a-worksheet)
12. [Right to left display: Set a worksheet into right to left display mode](#right-to-left-display-set-a-worksheet-into-right-to-left-display-mode)
13. [Autofitting Columns: Example of autofitting column widths](#autofitting-columns-example-of-autofitting-column-widths)
14. [Theme: Use a custom workbook theme](#theme-use-a-custom-workbook-theme)
15. [Theme: Use the Excel 2023/Aptos theme](#theme-use-the-excel-2023/aptos-theme)
16. [Insert images: Add images to a worksheet](#insert-images-add-images-to-a-worksheet)
17. [Insert images: Embedding an image in a cell](#insert-images-embedding-an-image-in-a-cell)
18. [Insert images: Inserting images to fit a cell](#insert-images-inserting-images-to-fit-a-cell)
19. [Adding a watermark: Adding a watermark to a worksheet by adding an image to the header](#adding-a-watermark-adding-a-watermark-to-a-worksheet-by-adding-an-image-to-the-header)
20. [Adding a watermark: Adding a watermark to a worksheet by adding a background image](#adding-a-watermark-adding-a-watermark-to-a-worksheet-by-adding-a-background-image)
21. [Chart: Simple: Simple getting started chart example](#chart-simple-simple-getting-started-chart-example)
22. [Chart: Area: Excel Area chart example](#chart-area-excel-area-chart-example)
23. [Chart: Bar: Excel Bar chart example](#chart-bar-excel-bar-chart-example)
24. [Chart: Column: Excel Column chart example](#chart-column-excel-column-chart-example)
25. [Chart: Line: Excel Line chart example](#chart-line-excel-line-chart-example)
26. [Chart: Scatter: Excel Scatter chart example](#chart-scatter-excel-scatter-chart-example)
27. [Chart: Pie: Excel Pie chart example](#chart-pie-excel-pie-chart-example)
28. [Chart: Doughnut: Excel Doughnut chart example](#chart-doughnut-excel-doughnut-chart-example)
29. [Chart: Radar: Excel Radar chart example](#chart-radar-excel-radar-chart-example)
30. [Chart: Stock: Excel Stock chart example](#chart-stock-excel-stock-chart-example)
31. [Chart: Using a secondary axis](#chart-using-a-secondary-axis)
32. [Chart: Create a combined chart](#chart-create-a-combined-chart)
33. [Chart: Create a combined pareto chart](#chart-create-a-combined-pareto-chart)
34. [Chart: Pattern Fill: Example of a chart with Pattern Fill](#chart-pattern-fill-example-of-a-chart-with-pattern-fill)
35. [Chart: Gradient Fill: Example of a chart with Gradient Fill](#chart-gradient-fill-example-of-a-chart-with-gradient-fill)
36. [Chart: Styles: Example of setting default chart styles](#chart-styles-example-of-setting-default-chart-styles)
37. [Chart: Chart data table](#chart-chart-data-table)
38. [Chart: Chart data tools](#chart-chart-data-tools)
39. [Chart: Clustered categories](#chart-clustered-categories)
40. [Chart: Gauge Chart](#chart-gauge-chart)
41. [Chart: Chartsheet](#chart-chartsheet)
42. [Grouped Rows: Create a grouped row outline](#grouped-rows-create-a-grouped-row-outline)
43. [Grouped Columns: Create a grouped column outline](#grouped-columns-create-a-grouped-column-outline)
44. [Textbox: Inserting Checkboxes in worksheets](#textbox-inserting-checkboxes-in-worksheets)
45. [Textbox: Inserting Textboxes in worksheets](#textbox-inserting-textboxes-in-worksheets)
46. [Textbox: Ignore Excel cell errors](#textbox-ignore-excel-cell-errors)
47. [Sparklines: simple example](#sparklines-simple-example)
48. [Sparklines: advanced example](#sparklines-advanced-example)
49. [Traits: Extending generic `write()` to handle user data types](#traits-extending-generic-write-to-handle-user-data-types)
50. [Macros: Adding macros to a workbook](#macros-adding-macros-to-a-workbook)
51. [Defined names: using user defined variable names in worksheets](#defined-names-using-user-defined-variable-names-in-worksheets)
52. [Cell Protection: Setting cell protection in a worksheet](#cell-protection-setting-cell-protection-in-a-worksheet)
53. [Document Properties: Setting document metadata properties for a workbook](#document-properties-setting-document-metadata-properties-for-a-workbook)
54. [Document Properties: Setting the Sensitivity Label](#document-properties-setting-the-sensitivity-label)
55. [Internal links: Creating a Table of Contents](#internal-links-creating-a-table-of-contents)
56. [Headers and Footers: Shows how to set headers and footers](#headers-and-footers-shows-how-to-set-headers-and-footers)
57. [Hyperlinks: Add hyperlinks to a worksheet](#hyperlinks-add-hyperlinks-to-a-worksheet)
58. [Freeze Panes: Example of setting freeze panes in worksheets](#freeze-panes-example-of-setting-freeze-panes-in-worksheets)
59. [Dynamic array formulas: Examples of dynamic arrays and formulas](#dynamic-array-formulas-examples-of-dynamic-arrays-and-formulas)
60. [Excel `LAMBDA()` function: Example of using the Excel 365 `LAMBDA()` function](#excel-lambda-function-example-of-using-the-excel-365-lambda-function)


# Hello World: Simple getting started example

Program to create a simple Hello World style Excel spreadsheet using the
`rust_xlsxwriter` library.

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
`rust_xlsxwriter` library. These are laid out on worksheets that correspond to the
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
            let text = format!("({col}, {row})");

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
[`Worksheet::merge_range()`].

[`Worksheet::merge_range()`]: crate::Worksheet::merge_range

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

An example of how to create autofilters with the `rust_xlsxwriter` library.

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


# Tables: Adding worksheet tables

Tables in Excel are a way of grouping a range of cells into a single entity
that has common formatting or that can be referenced from formulas. Tables
can have column headers, autofilters, total rows, column formulas and
different formatting styles.

The image below shows a default table in Excel with the default properties
shown in the ribbon bar.

<img src="https://rustxlsxwriter.github.io/images/table_intro.png">

A table is added to a worksheet via the [`Worksheet::add_table()`]method. The
headers and total row of a table should be configured via a [`Table`] struct but
the table data can be added via standard [`Worksheet::write()`]methods.

[`Table`]: crate::Table
[`Worksheet::write()`]: crate::Worksheet::write
[`Worksheet::add_table()`]: crate::Worksheet::add_table

## Some examples:

**Example 1.** Default table with no data.

<img src="https://rustxlsxwriter.github.io/images/app_tables1.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Default table with no data.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Create a new table.
    let table = Table::new();

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;
```


**Example 2.** Default table with data.

<img src="https://rustxlsxwriter.github.io/images/app_tables2.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Default table with data.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

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


**Example 3.** Table without default autofilter.

<img src="https://rustxlsxwriter.github.io/images/app_tables3.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table without default autofilter.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let table = Table::new().set_autofilter(false);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;
```


**Example 4.** Table without default header row.

<img src="https://rustxlsxwriter.github.io/images/app_tables4.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table without default header row.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let table = Table::new().set_header_row(false);

    // Add the table to the worksheet.
    worksheet.add_table(3, 1, 6, 5, &table)?;
```


**Example 5.** Default table with "First Column" and "Last Column" options.

<img src="https://rustxlsxwriter.github.io/images/app_tables5.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Default table with 'First Column' and 'Last Column' options.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let table = Table::new().set_first_column(true).set_last_column(true);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;
```


**Example 6.** Table with banded columns but without default banded rows.

<img src="https://rustxlsxwriter.github.io/images/app_tables6.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with banded columns but without default banded rows.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let table = Table::new().set_banded_rows(false).set_banded_columns(true);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;
```


**Example 7.** Table with user defined column headers.

<img src="https://rustxlsxwriter.github.io/images/app_tables7.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with user defined column headers.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let columns = vec![
        TableColumn::new().set_header("Product"),
        TableColumn::new().set_header("Quarter 1"),
        TableColumn::new().set_header("Quarter 2"),
        TableColumn::new().set_header("Quarter 3"),
        TableColumn::new().set_header("Quarter 4"),
    ];

    let table = Table::new().set_columns(&columns);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;
```


**Example 8.** Table with user defined column headers, and formulas.

<img src="https://rustxlsxwriter.github.io/images/app_tables8.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with user defined column headers, and formulas.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
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

    let table = Table::new().set_columns(&columns);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 6, &table)?;
```


**Example 9.** Table with totals row (but no caption or totals).

<img src="https://rustxlsxwriter.github.io/images/app_tables9.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with totals row (but no caption or totals).";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
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

    let table = Table::new().set_columns(&columns).set_total_row(true);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 7, 6, &table)?;
```


**Example 10.** Table with totals row with user captions and functions.

<img src="https://rustxlsxwriter.github.io/images/app_tables10.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with totals row with user captions and functions.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
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

    let table = Table::new().set_columns(&columns).set_total_row(true);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 7, 6, &table)?;
```


**Example 11.** Table with alternative Excel style.

<img src="https://rustxlsxwriter.github.io/images/app_tables11.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with alternative Excel style.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
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

    let table = Table::new()
        .set_columns(&columns)
        .set_total_row(true)
        .set_style(TableStyle::Light11);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 7, 6, &table)?;
```


**Example 12.** Table with Excel style removed.

<img src="https://rustxlsxwriter.github.io/images/app_tables12.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_tables.rs

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with Excel style removed.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
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

    let table = Table::new()
        .set_columns(&columns)
        .set_total_row(true)
        .set_style(TableStyle::None);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 7, 6, &table)?;
```


# Conditional Formatting: Adding conditional formatting to worksheets

Conditional formatting is a feature of Excel which allows you to apply a format
to a cell or a range of cells based on user defined rules. For example you might
apply rules like the following to highlight cells in different ranges.

<img src="https://rustxlsxwriter.github.io/images/conditional_format_dialog.png">

The examples below show how to use the various types of conditional formatting
with `rust_xlsxwriter`.

## Some examples:

**Example 1.** Cell conditional formatting. Cells with values >= 50 are in
light red. Values < 50 are in light green.

See [`ConditionalFormatCell`] for more details.

[`ConditionalFormatCell`]: crate::ConditionalFormatCell

<img src="https://rustxlsxwriter.github.io/images/conditional_formats1.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_conditional_formatting.rs

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
```


**Example 2.** Cell conditional formatting with between ranges. Values between
30 and 70 are in light red. Values outside that range are in light green.

See [`ConditionalFormatCell`] for more details.

[`ConditionalFormatCell`]: crate::ConditionalFormatCell

<img src="https://rustxlsxwriter.github.io/images/conditional_formats2.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_conditional_formatting.rs

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
```


**Example 3.** Duplicate and Unique conditional formats. Duplicate values are
in light red. Unique values are in light green.

See [`ConditionalFormatDuplicate`] for more details.

[`ConditionalFormatDuplicate`]: crate::ConditionalFormatDuplicate

<img src="https://rustxlsxwriter.github.io/images/conditional_formats3.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_conditional_formatting.rs

    // Write a conditional format over a range.
    let conditional_format = ConditionalFormatDuplicate::new().set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Invert the duplicate conditional format to show unique values in the
    // same range.
    let conditional_format = ConditionalFormatDuplicate::new()
        .invert()
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
```


**Example 4.** Above and Below Average conditional formats. Above average
values are in light red. Below average values are in light green.

See [`ConditionalFormatAverage`] for more details.

[`ConditionalFormatAverage`]: crate::ConditionalFormatAverage

<img src="https://rustxlsxwriter.github.io/images/conditional_formats4.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_conditional_formatting.rs

    // Write a conditional format over a range. The default criteria is Above Average.
    let conditional_format = ConditionalFormatAverage::new().set_format(&format1);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;

    // Write another conditional format over the same range.
    let conditional_format = ConditionalFormatAverage::new()
        .set_rule(ConditionalFormatAverageRule::BelowAverage)
        .set_format(&format2);

    worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
```


**Example 5.** Top and Bottom range conditional formats. Top 10 values are in
light red. Bottom 10 values are in light green.

See [`ConditionalFormatTop`] for more details.

[`ConditionalFormatTop`]: crate::ConditionalFormatTop

<img src="https://rustxlsxwriter.github.io/images/conditional_formats5.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_conditional_formatting.rs

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
```


**Example 6.** Cell conditional formatting in non-contiguous range. Cells with
values >= 50 are in light red. Values < 50 are in light green. Non-contiguous
ranges.

<img src="https://rustxlsxwriter.github.io/images/conditional_formats6.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_conditional_formatting.rs

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
```


**Example 7.** Formula conditional formatting. Even numbered cells are in
light green. Odd numbered cells are in light red.

See [`ConditionalFormatFormula`] for more details.

[`ConditionalFormatFormula`]: crate::ConditionalFormatFormula

<img src="https://rustxlsxwriter.github.io/images/conditional_formats7.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_conditional_formatting.rs

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
```


**Example 8.** Text style conditional formats. Column A shows words that
contain the sub-word 'rust'. Column C shows words that start/end with 't'

See [`ConditionalFormatText`] for more details.

[`ConditionalFormatText`]: crate::ConditionalFormatText

<img src="https://rustxlsxwriter.github.io/images/conditional_formats8.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_conditional_formatting.rs

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
```


**Example 9.** Examples of 2 color scale conditional formats.

See [`ConditionalFormat2ColorScale`] for more details.

[`ConditionalFormat2ColorScale`]: crate::ConditionalFormat2ColorScale

<img src="https://rustxlsxwriter.github.io/images/conditional_formats9.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_conditional_formatting.rs

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
```


**Example 10.** Examples of 3 color scale conditional formats.

See [`ConditionalFormat3ColorScale`] for more details.

[`ConditionalFormat3ColorScale`]: crate::ConditionalFormat3ColorScale

<img src="https://rustxlsxwriter.github.io/images/conditional_formats10.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_conditional_formatting.rs

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
```


**Example 11.** Examples of data bars.

See [`ConditionalFormatDataBar`] for more details.

[`ConditionalFormatDataBar`]: crate::ConditionalFormatDataBar

<img src="https://rustxlsxwriter.github.io/images/conditional_formats11.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_conditional_formatting.rs

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
```


**Example 12.** Examples of icon style conditional formats.


See [`ConditionalFormatIconSet`] for more details.

[`ConditionalFormatIconSet`]: crate::ConditionalFormatIconSet

<img src="https://rustxlsxwriter.github.io/images/conditional_formats12.png">

Code to generate the above example:

```ignore
    // Code snippet from examples/app_conditional_formatting.rs

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
```


# Data Validation: Add cell validation and dropdowns

Example of how to add data validation and dropdown lists using the
`rust_xlsxwriter` library.

Data validation is a feature of Excel which allows you to restrict the data that
a user enters in a cell and to display help and warning messages. It also allows
you to restrict input to values in a drop down list.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_data_validation.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_data_validation.rs

use rust_xlsxwriter::{
    DataValidation, DataValidationErrorStyle, DataValidationRule, ExcelDateTime, Format,
    FormatAlign, FormatBorder, Formula, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add a format for the header cells.
    let header_format = Format::new()
        .set_background_color("C6EFCE")
        .set_border(FormatBorder::Thin)
        .set_bold()
        .set_indent(1)
        .set_text_wrap()
        .set_align(FormatAlign::VerticalCenter);

    // Set up layout of the worksheet.
    worksheet.set_column_width(0, 68)?;
    worksheet.set_column_width(1, 15)?;
    worksheet.set_column_width(3, 15)?;
    worksheet.set_row_height(0, 36)?;

    // Write the header cells and some data that will be used in the examples.
    let heading1 = "Some examples of data validations";
    let heading2 = "Enter values in this column";
    let heading3 = "Sample Data";

    worksheet.write_with_format(0, 0, heading1, &header_format)?;
    worksheet.write_with_format(0, 1, heading2, &header_format)?;
    worksheet.write_with_format(0, 3, heading3, &header_format)?;

    worksheet.write(2, 3, "Integers")?;
    worksheet.write(2, 4, 1)?;
    worksheet.write(2, 5, 10)?;

    worksheet.write_row(3, 3, ["List data", "open", "high", "close"])?;

    worksheet.write(4, 3, "Formula")?;
    worksheet.write(4, 4, Formula::new("=AND(F5=50,G5=60)"))?;
    worksheet.write(4, 5, 50)?;
    worksheet.write(4, 6, 60)?;

    // -----------------------------------------------------------------------
    // Example 1. Limiting input to an integer in a fixed range.
    // -----------------------------------------------------------------------
    let text = "Enter an integer between 1 and 10";
    worksheet.write(2, 0, text)?;

    let data_validation =
        DataValidation::new().allow_whole_number(DataValidationRule::Between(1, 10));

    worksheet.add_data_validation(2, 1, 2, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 2. Limiting input to an integer outside a fixed range.
    // -----------------------------------------------------------------------
    let text = "Enter an integer that is not between 1 and 10 (using cell references)";
    worksheet.write(4, 0, text)?;

    let data_validation = DataValidation::new()
        .allow_whole_number_formula(DataValidationRule::NotBetween("=E3".into(), "=F3".into()));

    worksheet.add_data_validation(4, 1, 4, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 3. Limiting input to an integer greater than a fixed value.
    // -----------------------------------------------------------------------
    let text = "Enter an integer greater than 0";
    worksheet.write(6, 0, text)?;

    let data_validation =
        DataValidation::new().allow_whole_number(DataValidationRule::GreaterThan(0));

    worksheet.add_data_validation(6, 1, 6, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 4. Limiting input to an integer less than a fixed value.
    // -----------------------------------------------------------------------
    let text = "Enter an integer less than 10";
    worksheet.write(8, 0, text)?;

    let data_validation =
        DataValidation::new().allow_whole_number(DataValidationRule::LessThan(10));

    worksheet.add_data_validation(8, 1, 8, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 5. Limiting input to a decimal in a fixed range.
    // -----------------------------------------------------------------------
    let text = "Enter a decimal between 0.1 and 0.5";
    worksheet.write(10, 0, text)?;

    let data_validation =
        DataValidation::new().allow_decimal_number(DataValidationRule::Between(0.1, 0.5));

    worksheet.add_data_validation(10, 1, 10, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 6. Limiting input to a value in a dropdown list.
    // -----------------------------------------------------------------------
    let text = "Select a value from a drop down list";
    worksheet.write(12, 0, text)?;

    let data_validation = DataValidation::new().allow_list_strings(&["open", "high", "close"])?;

    worksheet.add_data_validation(12, 1, 12, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 7. Limiting input to a value in a dropdown list.
    // -----------------------------------------------------------------------
    let text = "Select a value from a drop down list (using a cell range)";
    worksheet.write(14, 0, text)?;

    let data_validation = DataValidation::new().allow_list_formula("=$E$4:$G$4".into());

    worksheet.add_data_validation(14, 1, 14, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 8. Limiting input to a date in a fixed range.
    // -----------------------------------------------------------------------
    let text = "Enter a date between 1/1/2025 and 12/12/2025";
    worksheet.write(16, 0, text)?;

    let data_validation = DataValidation::new().allow_date(DataValidationRule::Between(
        ExcelDateTime::parse_from_str("2025-01-01")?,
        ExcelDateTime::parse_from_str("2025-12-12")?,
    ));

    worksheet.add_data_validation(16, 1, 16, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 9. Limiting input to a time in a fixed range.
    // -----------------------------------------------------------------------
    let text = "Enter a time between 6:00 and 12:00";
    worksheet.write(18, 0, text)?;

    let data_validation = DataValidation::new().allow_time(DataValidationRule::Between(
        ExcelDateTime::parse_from_str("6:00")?,
        ExcelDateTime::parse_from_str("12:00")?,
    ));

    worksheet.add_data_validation(18, 1, 18, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 10. Limiting input to a string greater than a fixed length.
    // -----------------------------------------------------------------------
    let text = "Enter a string longer than 3 characters";
    worksheet.write(20, 0, text)?;

    let data_validation =
        DataValidation::new().allow_text_length(DataValidationRule::GreaterThan(3));

    worksheet.add_data_validation(20, 1, 20, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 11. Limiting input based on a formula.
    // -----------------------------------------------------------------------
    let text = "Enter a value if the following is true '=AND(F5=50,G5=60)'";
    worksheet.write(22, 0, text)?;

    let data_validation = DataValidation::new().allow_custom("=AND(F5=50,G5=60)".into());

    worksheet.add_data_validation(22, 1, 22, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 12. Displaying and modifying data validation messages.
    // -----------------------------------------------------------------------
    let text = "Displays a message when you select the cell";
    worksheet.write(24, 0, text)?;

    let data_validation = DataValidation::new()
        .allow_whole_number(DataValidationRule::Between(1, 100))
        .set_input_title("Enter an integer:")?
        .set_input_message("between 1 and 100")?;

    worksheet.add_data_validation(24, 1, 24, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 13. Displaying and modifying data validation messages.
    // -----------------------------------------------------------------------
    let text = "Display a custom error message when integer isn't between 1 and 100";
    worksheet.write(26, 0, text)?;

    let data_validation = DataValidation::new()
        .allow_whole_number(DataValidationRule::Between(1, 100))
        .set_input_title("Enter an integer:")?
        .set_input_message("between 1 and 100")?
        .set_error_title("Input value is not valid!")?
        .set_error_message("It should be an integer between 1 and 100")?;

    worksheet.add_data_validation(26, 1, 26, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Example 14. Displaying and modifying data validation messages.
    // -----------------------------------------------------------------------
    let text = "Display a custom info message when integer isn't between 1 and 100";
    worksheet.write(28, 0, text)?;

    let data_validation = DataValidation::new()
        .allow_whole_number(DataValidationRule::Between(1, 100))
        .set_input_title("Enter an integer:")?
        .set_input_message("between 1 and 100")?
        .set_error_title("Input value is not valid!")?
        .set_error_message("It should be an integer between 1 and 100")?
        .set_error_style(DataValidationErrorStyle::Information);

    worksheet.add_data_validation(28, 1, 28, 1, &data_validation)?;

    // -----------------------------------------------------------------------
    // Save and close the file.
    // -----------------------------------------------------------------------
    workbook.save("data_validation.xlsx")?;

    Ok(())
}
```


# Notes: Adding notes to worksheet cells

An example of writing cell Notes to a worksheet.

A Note is a post-it style message that is revealed when the user mouses
over a worksheet cell. The presence of a Note is indicated by a small
red triangle in the upper right-hand corner of the cell.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_notes.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_notes.rs

use rust_xlsxwriter::{Note, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Widen the first column for clarity.
    worksheet.set_column_width(0, 16)?;

    // Write some data.
    let party_items = [
        "Invitations",
        "Doors",
        "Flowers",
        "Champagne",
        "Menu",
        "Peter",
    ];
    worksheet.write_column(0, 0, party_items)?;

    // Create a new worksheet Note.
    let note = Note::new("I will get the flowers myself").set_author("Clarissa Dalloway");

    // Add the note to a cell.
    worksheet.insert_note(2, 0, &note)?;

    // Save the file to disk.
    workbook.save("notes.xlsx")?;

    Ok(())
}
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
    worksheet1.write(0, 0, "نص عربي / English text")?;
    worksheet1.write_with_format(1, 0, "نص عربي / English text", &format_left_to_right)?;
    worksheet1.write_with_format(2, 0, "نص عربي / English text", &format_right_to_left)?;

    // Add a worksheet and change it to right to left direction.
    let worksheet2 = workbook.add_worksheet();
    worksheet2.set_right_to_left(true);

    // Make the column wider for clarity.
    worksheet2.set_column_width(0, 25)?;

    // Right to left direction:    ... | C1 | B1 | A1 |
    worksheet2.write(0, 0, "نص عربي / English text")?;
    worksheet2.write_with_format(1, 0, "نص عربي / English text", &format_left_to_right)?;
    worksheet2.write_with_format(2, 0, "نص عربي / English text", &format_right_to_left)?;

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


# Theme: Use a custom workbook theme

An example of setting the default theme for a workbook to a user supplied custom
theme using the `rust_xlsxwriter` library. The theme xml file is extracted from
an Excel xlsx file. Note that the default font has changed to "Arial (body) 11".

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_theme_custom.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_theme_custom.rs

use rust_xlsxwriter::{FontScheme, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a custom theme to the workbook.
    workbook.use_custom_theme("tests/input/themes/technic.xml")?;

    // Create a new default format to match the custom theme. Note, that the
    // scheme is set to "Body" to indicate that the font is part of the theme.
    let format = Format::new()
        .set_font_name("Arial")
        .set_font_size(11)
        .set_font_scheme(FontScheme::Body);

    // Add the default format for the workbook.
    workbook.set_default_format(&format, 19, 72)?;

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some text to demonstrate the changed theme.
    worksheet.write(0, 0, "Hello")?;

    // Save the workbook to disk.
    workbook.save("theme_custom.xlsx")?;

    Ok(())
}
```


# Theme: Use the Excel 2023/Aptos theme

An example of changing the default theme for a workbook using the
`rust_xlsxwriter` library. The example uses the Excel 2023 Office/Aptos theme.
Note that the default font has changed to "Aptos Narrow (body) 11".

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_theme_excel_2023.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_theme_excel_2023.rs

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Use the Excel 2023 Office/Aptos theme in the workbook.
    workbook.use_excel_2023_theme()?;

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some text to demonstrate the changed theme.
    worksheet.write(0, 0, "Hello")?;

    // Save the workbook to disk.
    workbook.save("theme_excel_2023.xlsx")?;

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
    image = image.set_scale_width(0.75).set_scale_height(0.75);
    worksheet.insert_image(15, 1, &image)?;

    // Save the file to disk.
    workbook.save("images.xlsx")?;

    Ok(())
}
```


# Insert images: Embedding an image in a cell

An example of embedding images into a worksheet cells using `rust_xlsxwriter`.
This image scales to size of the cell and moves with it.

This approach can be useful if you are building up a spreadsheet of products
with a column of images for each product.

This is the equivalent of Excel's menu option to insert an image using the
option to "Place in Cell" which is only available in Excel 365 versions from
2023 onwards. For older versions of Excel a `#VALUE!` error is displayed.


**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/embedded_images.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_embedded_images.rs

use rust_xlsxwriter::{Format, FormatAlign, Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create some formats to use in the example.
    let vertical_center = Format::new().set_align(FormatAlign::VerticalCenter);
    let center = Format::new().set_align(FormatAlign::Center);

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a new image object.
    let image = Image::new("examples/rust_logo.png")?;

    // Widen the first column to make the captions clearer.
    worksheet.set_column_width(0, 30)?;

    // Change cell widths/heights to demonstrate the image differences.
    worksheet.set_column_width(1, 14)?;
    worksheet.set_row_height(1, 60)?;
    worksheet.set_row_height(3, 60)?;
    worksheet.set_row_height(5, 90)?;

    // Embed an image in a cell. The height and width scale automatically.
    worksheet.write_with_format(1, 0, "Embed image in cell:", &vertical_center)?;
    worksheet.embed_image(1, 1, &image)?;

    // Embed and center an image in a cell.
    worksheet.write_with_format(3, 0, "Embed and center image:", &vertical_center)?;
    worksheet.embed_image_with_format(3, 1, &image, &center)?;

    // Embed an image in a larger cell.
    worksheet.write_with_format(5, 0, "Embed image in larger cell:", &vertical_center)?;
    worksheet.embed_image(5, 1, &image)?;

    // Save the file to disk.
    workbook.save("embedded_images.xlsx")?;

    Ok(())
}
```


# Insert images: Inserting images to fit a cell

An example of inserting images into a worksheet using `rust_xlsxwriter`so that
they are scaled to a cell. This approach can be useful if you are building up a
spreadsheet of products with a column of images for each product.

See the [Embedding images in cells](#insert-images-embedding-an-image-in-a-cell) example that
shows a better approach for newer versions of Excel.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_images_fit_to_cell.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_images_fit_to_cell.rs

//! for newer versions of Excel.

use rust_xlsxwriter::{Format, FormatAlign, Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    let center = Format::new().set_align(FormatAlign::VerticalCenter);

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Widen the first column to make the text clearer.
    worksheet.set_column_width_pixels(0, 250)?;

    // Set larger cells to accommodate the images.
    worksheet.set_column_width_pixels(1, 200)?;
    worksheet.set_row_height_pixels(0, 140)?;
    worksheet.set_row_height_pixels(2, 140)?;
    worksheet.set_row_height_pixels(4, 140)?;
    worksheet.set_row_height_pixels(6, 140)?;

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

    // Insert the image and scale it to the cell while keeping it centered. This
    // also maintains the aspect ratio.
    worksheet.write_with_format(6, 0, "Image scaled and centered:", &center)?;
    worksheet.insert_image_fit_to_cell_centered(6, 1, &image)?;

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


# Adding a watermark: Adding a watermark to a worksheet by adding a background image

An example of adding a background image to a worksheet. In this case it is used as a watermark.

See also the previous example where a watermark is created by putting an image
in the worksheet header.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_background_image.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_background_image.rs

use rust_xlsxwriter::{Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // The image may not be visible unless the view is large.
    worksheet.write(0, 0, "Scroll down and right to see the background image")?;

    // Create a new image object.
    let image = Image::new("examples/watermark.png")?;

    // Insert the background image.
    worksheet.insert_background_image(&image);

    // Save the file to disk.
    workbook.save("background_image.xlsx")?;

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


# Chart: Stock: Excel Stock chart example

Example of creating Excel Stock charts.

Note, Volume variants of the Excel stock charts aren't currently supported but
will be in a future release.

**Image of the output file:**

Chart 1 in the following example is an example of a High-Low-Close Stock chart.

To create a chart similar to a default Excel High-Low-Close Stock chart
you need to do the following steps:

1. Create a `Stock` type chart.
2. Add 3 series for High, Low and Close, in that order.
3. Hide the default lines in all 3 series.
4. Hide the default markers for the High and Low series.
5. Set a dash marker for the Close series.
6. Turn on the chart High-Low bars.
7. Format any other chart or axis property you need.


<img src="https://rustxlsxwriter.github.io/images/chart_stock1.png">

Chart 2 in the following example is an example of an Open-High-Low-Close Stock chart.

To create a chart similar to a default Excel Open-High-Low-Close Stock
chart you need to do the following steps:

1. Create a `Stock` type chart.
2. Add 4 series for Open, High, Low and Close, in that order.
3. Hide the default lines in all 4 series.
4. Turn on the chart High-Low bars.
5. Turn on the chart Up-Down bars.
6. Format any other chart or axis property you need.


<img src="https://rustxlsxwriter.github.io/images/chart_stock2.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_stock.rs

use rust_xlsxwriter::{
    Chart, ChartFormat, ChartLine, ChartMarker, ChartMarkerType, ChartSolidFill, ChartType,
    ExcelDateTime, Format, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Create some formatting to use for the worksheet data.
    let bold = Format::new().set_bold();
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let money_format = Format::new().set_num_format("[$$-en-US]#,##0.00");

    // -----------------------------------------------------------------------
    // Create some simulated stock data for the chart.
    //
    let dates = [
        "2024-05-01",
        "2024-05-02",
        "2024-05-03",
        "2024-05-04",
        "2024-05-05",
        "2024-05-06",
        "2024-05-07",
        "2024-05-08",
        "2024-05-09",
        "2024-05-10",
    ];

    // Map the string dates to ExcelDateTime objects, while capturing any
    // potential conversion errors.
    let dates: Result<Vec<ExcelDateTime>, XlsxError> = dates
        .into_iter()
        .map(ExcelDateTime::parse_from_str)
        .collect();
    let dates = dates?;

    let open_data = [
        35.00, 41.53, 43.33, 46.73, 49.50, 53.29, 59.56, 30.18, 25.83, 20.65,
    ];

    let high_data = [
        44.12, 45.98, 46.99, 50.40, 54.99, 60.32, 30.45, 26.51, 23.02, 30.10,
    ];

    let low_low = [
        32.59, 38.51, 40.02, 45.60, 47.17, 52.02, 59.11, 28.97, 25.06, 18.25,
    ];

    let close_data = [
        41.53, 43.33, 46.73, 49.50, 53.29, 59.56, 30.18, 25.83, 20.65, 28.00,
    ];

    // -----------------------------------------------------------------------
    // Define variables so that the chart can change dynamically with the data.
    //
    let header_row = 0;
    let start_row = header_row + 1;
    let end_row = start_row + (dates.len() as u32) - 1;
    let date_col = 0;
    let open_col = date_col + 1;
    let high_col = date_col + 2;
    let low_col = date_col + 3;
    let close_col = date_col + 4;

    // -----------------------------------------------------------------------
    // Write the data to the worksheet, with formatting.
    //
    worksheet.write_with_format(header_row, date_col, "Date", &bold)?;
    worksheet.write_with_format(header_row, open_col, "Open", &bold)?;
    worksheet.write_with_format(header_row, high_col, "High", &bold)?;
    worksheet.write_with_format(header_row, low_col, "Low", &bold)?;
    worksheet.write_with_format(header_row, close_col, "Close", &bold)?;

    worksheet.write_column_with_format(start_row, date_col, dates, &date_format)?;
    worksheet.write_column_with_format(start_row, open_col, open_data, &money_format)?;
    worksheet.write_column_with_format(start_row, high_col, high_data, &money_format)?;
    worksheet.write_column_with_format(start_row, low_col, low_low, &money_format)?;
    worksheet.write_column_with_format(start_row, close_col, close_data, &money_format)?;

    // Change the width of the date column, for clarity.
    worksheet.set_column_width(date_col, 11)?;

    // -----------------------------------------------------------------------
    // Create a new High-Low-Close Stock chart.
    //
    // To create a chart similar to a default Excel High-Low-Close Stock chart
    // you need to do the following steps:
    //
    // 1. Create a `Stock` type chart.
    // 2. Add 3 series for High, Low and Close, in that order.
    // 3. Hide the default lines in all 3 series.
    // 4. Hide the default markers for the High and Low series.
    // 5. Set a dash marker for the Close series.
    // 6. Turn on the chart High-Low bars.
    // 7. Format any other chart or axis property you need.
    //
    let mut chart = Chart::new(ChartType::Stock);

    // Add the High series.
    chart
        .add_series()
        .set_categories(("Sheet1", start_row, date_col, end_row, date_col))
        .set_values(("Sheet1", start_row, high_col, end_row, high_col))
        .set_format(ChartLine::new().set_hidden(true))
        .set_marker(ChartMarker::new().set_none());

    // Add the Low series.
    chart
        .add_series()
        .set_categories(("Sheet1", start_row, date_col, end_row, date_col))
        .set_values(("Sheet1", start_row, low_col, end_row, low_col))
        .set_format(ChartLine::new().set_hidden(true))
        .set_marker(ChartMarker::new().set_none());

    // Add the Close series.
    chart
        .add_series()
        .set_categories(("Sheet1", start_row, date_col, end_row, date_col))
        .set_values(("Sheet1", start_row, close_col, end_row, close_col))
        .set_format(ChartLine::new().set_hidden(true))
        .set_marker(
            ChartMarker::new()
                .set_type(ChartMarkerType::LongDash)
                .set_size(10)
                .set_format(
                    ChartFormat::new()
                        .set_border(ChartLine::new().set_color("#000000"))
                        .set_solid_fill(ChartSolidFill::new().set_color("#000000")),
                ),
        );

    // Set the High-Low lines.
    chart.set_high_low_lines(true);

    // Add a chart title and some axis labels.
    chart.title().set_name("Stock: High - Low - Close");
    chart.x_axis().set_name("Date");
    chart.y_axis().set_name("Stock Price");

    // Format the price axis number format.
    chart.y_axis().set_num_format("[$$-en-US]#,##0");

    // Turn off the chart legend.
    chart.legend().set_hidden();

    // Insert the chart into the worksheet.
    worksheet.insert_chart_with_offset(start_row, close_col + 1, &chart, 20, 10)?;

    // -----------------------------------------------------------------------
    // Create a new Open-High-Low-Close Stock chart.
    //
    // To create a chart similar to a default Excel Open-High-Low-Close Stock
    // chart you need to do the following steps:
    //
    // 1. Create a `Stock` type chart.
    // 2. Add 4 series for Open, High, Low and Close, in that order.
    // 3. Hide the default lines in all 4 series.
    // 4. Turn on the chart High-Low bars.
    // 5. Turn on the chart Up-Down bars.
    // 6. Format any other chart or axis property you need.
    //
    let mut chart = Chart::new(ChartType::Stock);

    // Add the Open series.
    chart
        .add_series()
        .set_categories(("Sheet1", start_row, date_col, end_row, date_col))
        .set_values(("Sheet1", start_row, open_col, end_row, open_col))
        .set_format(ChartLine::new().set_hidden(true))
        .set_marker(ChartMarker::new().set_none());

    // Add the High series.
    chart
        .add_series()
        .set_categories(("Sheet1", start_row, date_col, end_row, date_col))
        .set_values(("Sheet1", start_row, high_col, end_row, high_col))
        .set_format(ChartLine::new().set_hidden(true))
        .set_marker(ChartMarker::new().set_none());

    // Add the Low series.
    chart
        .add_series()
        .set_categories(("Sheet1", start_row, date_col, end_row, date_col))
        .set_values(("Sheet1", start_row, low_col, end_row, low_col))
        .set_format(ChartLine::new().set_hidden(true))
        .set_marker(ChartMarker::new().set_none());

    // Add the Close series.
    chart
        .add_series()
        .set_categories(("Sheet1", start_row, date_col, end_row, date_col))
        .set_values(("Sheet1", start_row, close_col, end_row, close_col))
        .set_format(ChartLine::new().set_hidden(true))
        .set_marker(ChartMarker::new().set_none());

    // Set the High-Low lines.
    chart.set_high_low_lines(true);

    // Turn on and format the Up-Down bars.
    chart.set_up_down_bars(true);
    chart.set_up_bar_format(ChartSolidFill::new().set_color("#009933"));
    chart.set_down_bar_format(ChartSolidFill::new().set_color("#FF5050"));

    // Add a chart title and some axis labels.
    chart.title().set_name("Stock: Open - High - Low - Close");
    chart.x_axis().set_name("Date");
    chart.y_axis().set_name("Stock Price");

    // Format the price axis number format.
    chart.y_axis().set_num_format("[$$-en-US]#,##0");

    // Turn off the chart legend.
    chart.legend().set_hidden();

    // Insert the chart into the worksheet.
    worksheet.insert_chart_with_offset(start_row + 16, close_col + 1, &chart, 20, 10)?;

    // -----------------------------------------------------------------------
    // Save and close the file.
    // -----------------------------------------------------------------------
    workbook.save("chart_stock.xlsx")?;

    Ok(())
}
```


# Chart: Using a secondary axis


Example of creating an Excel Line chart with a secondary axis by setting the
[`ChartSeries::set_secondary_axis()`] property for one of more series in the
chart.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_chart_secondary_axis.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_secondary_axis.rs

use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Aliens", &bold)?;
    worksheet.write_with_format(0, 1, "Humans", &bold)?;
    worksheet.write_column(1, 0, [2, 3, 4, 5, 6, 7])?;
    worksheet.write_column(1, 1, [10, 40, 50, 20, 10, 50])?;

    // Create a new line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure a series with a secondary axis.
    chart
        .add_series()
        .set_name("Sheet1!$A$1")
        .set_values("Sheet1!$A$2:$A$7")
        .set_secondary_axis(true);

    // Configure another series that defaults to the primary axis.
    chart
        .add_series()
        .set_name("Sheet1!$B$1")
        .set_values("Sheet1!$B$2:$B$7");

    // Add a chart title and some axis labels.
    chart.title().set_name("Survey results");
    chart.x_axis().set_name("Days");
    chart.y_axis().set_name("Population");
    chart.y2_axis().set_name("Laser wounds");
    chart.y_axis().set_major_gridlines(false);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 3, &chart, 25, 10)?;

    workbook.save("chart_secondary_axis.xlsx")?;

    Ok(())
}
```

[`ChartSeries::set_secondary_axis()`]: crate::ChartSeries::set_secondary_axis


In general secondary axes are used for displaying different Y values for the
same category range. However it is also possible to display a secondary X axis
for series that use a different category range. See the example below.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/chart_series_set_secondary_axis2.png">


**Code to generate the output file:**

```rust
// Sample code from examples/doc_chart_series_set_secondary_axis2.rs

use rust_xlsxwriter::{
    Chart, ChartAxisCrossing, ChartAxisLabelPosition, ChartLegendPosition, ChartType, Workbook,
    XlsxError,
};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_column(0, 0, [1, 2, 3, 4, 5])?;
    worksheet.write_column(0, 1, [10, 40, 50, 20, 10])?;
    worksheet.write_column(0, 2, [1, 2, 3, 4, 5, 6, 7])?;
    worksheet.write_column(0, 3, [30, 10, 20, 40, 30, 10, 20])?;

    // Create a new line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure a series that defaults to the primary axis.
    chart
        .add_series()
        .set_categories(("Sheet1", 0, 0, 4, 0))
        .set_values(("Sheet1", 0, 1, 4, 1));

    // Configure another series with a secondary axis. Note that the category
    // range is different to the primary axes series.
    chart
        .add_series()
        .set_categories(("Sheet1", 0, 2, 6, 2))
        .set_values(("Sheet1", 0, 3, 6, 3))
        .set_secondary_axis(true);

    // Make the secondary X axis visible (it is hidden by default) and also
    // position the labels so they are next to the axis and therefore visible.
    chart
        .x2_axis()
        .set_hidden(false)
        .set_label_position(ChartAxisLabelPosition::NextTo);

    // Set the X2 axis to cross the Y2 axis at the max value so it appears at
    // the top of the chart.
    chart.y2_axis().set_crossing(ChartAxisCrossing::Max);

    // Add some axis labels.
    chart.x_axis().set_name("X axis");
    chart.y_axis().set_name("Y axis");
    chart.x2_axis().set_name("X2 axis");
    chart.y2_axis().set_name("Y2 axis");

    // Move the legend to the bottom for clarity.
    chart.legend().set_position(ChartLegendPosition::Bottom);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(0, 4, &chart, 5, 5)?;

    workbook.save("chart.xlsx")?;

    Ok(())
}
```




# Chart: Create a combined chart

Example of creating combined Excel charts from two different chart types.

**Image of the output file:**

In the first example we create a combined column and line chart that share the
same X and Y axes:

<img src="https://rustxlsxwriter.github.io/images/app_chart_combined1.png">

In the second example we create a similar combined column and line chart except
that the secondary chart has a secondary Y axis:

<img src="https://rustxlsxwriter.github.io/images/app_chart_combined2.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_combined.rs

use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    let headings = ["Number", "Sample", "Target"];
    worksheet.write_row_with_format(0, 0, headings, &bold)?;

    let data = [
        [2, 3, 4, 5, 6, 7],
        [10, 40, 50, 20, 10, 50],
        [30, 60, 70, 50, 40, 30],
    ];
    worksheet.write_column_matrix(1, 0, data)?;

    // -----------------------------------------------------------------------
    // In the first example we will create a combined column and line chart.
    // The charts will share the same X and Y axes.
    // -----------------------------------------------------------------------
    let mut column_chart = Chart::new(ChartType::Column);

    // Configure the data series for the primary chart.
    column_chart
        .add_series()
        .set_name("Sheet1!$B$1")
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7");

    // Create a new line chart. This will use this as the secondary chart.
    let mut line_chart = Chart::new(ChartType::Line);

    // Configure the data series for the secondary chart.
    line_chart
        .add_series()
        .set_name("Sheet1!$C$1")
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Combine the charts.
    column_chart.combine(&line_chart);

    // Add a chart title and some axis labels. Note, this is done via the
    // primary chart.
    column_chart
        .title()
        .set_name("Combined chart with same Y axis");
    column_chart.x_axis().set_name("Test number");
    column_chart.y_axis().set_name("Sample length (mm)");

    // Add the primary chart to the worksheet.
    worksheet.insert_chart(1, 4, &column_chart)?;

    // -----------------------------------------------------------------------
    // In the second example we will create a similar combined column and line
    // chart except that the secondary chart will have a secondary Y axis.
    // -----------------------------------------------------------------------
    let mut column_chart = Chart::new(ChartType::Column);

    // Configure the data series for the primary chart.
    column_chart
        .add_series()
        .set_name("Sheet1!$B$1")
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7");

    // Create a new line chart. This will use this as the secondary chart.
    let mut line_chart = Chart::new(ChartType::Line);

    // Configure the data series for the secondary chart.
    line_chart
        .add_series()
        .set_name("Sheet1!$C$1")
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_secondary_axis(true);

    // Combine the charts.
    column_chart.combine(&line_chart);

    // Configure the data series for the secondary chart. We also set a
    // secondary Y axis via (y2_axis). This is the only difference between
    // this and the first example, apart from the axis label below.
    column_chart
        .title()
        .set_name("Combine chart with secondary Y axis");
    column_chart.x_axis().set_name("Test number");
    column_chart.y_axis().set_name("Sample length (mm)");

    // Note: the y2 properties are set via the primary chart.
    column_chart.y2_axis().set_name("Target length (mm)");

    // Add the primary chart to the worksheet.
    worksheet.insert_chart(17, 4, &column_chart)?;

    // Save the file to disk.
    workbook.save("chart_combined.xlsx")?;

    Ok(())
}
```


# Chart: Create a combined pareto chart

Example of creating a Pareto chart with a secondary chart and axis.

A Pareto chart is a type of chart that combines a Column/Histogram chart and a
Chart. Individual values are represented in descending order by the columns and
the cumulative total is represented by the line approaching 100% on a second
axis.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_chart_pareto.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_pareto.rs

use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Formats used in the workbook.
    let bold = Format::new().set_bold();
    let percent_format = Format::new().set_num_format("0%");

    // Add the worksheet data that the charts will refer to.
    let headings = ["Reason", "Number", "Percentage"];

    let reasons = [
        "Traffic",
        "Child care",
        "Public Transport",
        "Weather",
        "Overslept",
        "Emergency",
    ];

    let numbers = [60, 40, 20, 15, 10, 5];
    let percents = [0.440, 0.667, 0.800, 0.900, 0.967, 1.00];

    worksheet.write_row_with_format(0, 0, headings, &bold)?;
    worksheet.write_column(1, 0, reasons)?;
    worksheet.write_column(1, 1, numbers)?;
    worksheet.write_column_with_format(1, 2, percents, &percent_format)?;

    // Widen the columns for visibility.
    worksheet.set_column_width(0, 15)?;
    worksheet.set_column_width(1, 10)?;
    worksheet.set_column_width(2, 10)?;

    //
    // Create a new Column chart. This will be the primary chart.
    //
    let mut column_chart = Chart::new(ChartType::Column);

    // Configure a series on the primary axis.
    column_chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7");

    // Add a chart title.
    column_chart.title().set_name("Reasons for lateness");

    // Turn off the chart legend.
    column_chart.legend().set_hidden();

    // Set the  name and scale of the Y axes. Note, the secondary axis is set
    // from the primary chart.
    column_chart
        .y_axis()
        .set_name("Respondents (number)")
        .set_min(0)
        .set_max(120);

    column_chart.y2_axis().set_max(1);

    //
    // Create a new Line chart. This will be the secondary chart.
    //
    let mut line_chart = Chart::new(ChartType::Line);

    // Add a series on the secondary axis.
    line_chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_secondary_axis(true);

    // Combine the charts.
    column_chart.combine(&line_chart);

    // Add the chart to the worksheet.
    worksheet.insert_chart(1, 5, &column_chart)?;

    workbook.save("chart_pareto.xlsx")?;

    Ok(())
}
```


# Chart: Pattern Fill: Example of a chart with Pattern Fill

A example of creating column charts with fill patterns using the [`ChartFormat`]
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


# Chart: Gradient Fill: Example of a chart with Gradient Fill

A example of creating column charts with fill gradients using the
[`ChartFormat`] and [`ChartGradientFill`] structs.


**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/chart_gradient_fill.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_gradient.rs

use rust_xlsxwriter::{
    Chart, ChartGradientFill, ChartGradientStop, ChartType, Format, Workbook, XlsxError,
};

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

    // Create a new column chart.
    let mut chart = Chart::new(ChartType::Column);

    //
    // Create a gradient profile to the first series.
    //
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1")
        .set_format(ChartGradientFill::new().set_gradient_stops(&[
            ChartGradientStop::new("#963735", 0),
            ChartGradientStop::new("#F1DCDB", 100),
        ]));

    //
    // Create a gradient profile to the second series.
    //
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1")
        .set_format(ChartGradientFill::new().set_gradient_stops(&[
            ChartGradientStop::new("#E36C0A", 0),
            ChartGradientStop::new("#FCEADA", 100),
        ]));

    //
    // Create a gradient profile and add it to chart plot area.
    //
    chart
        .plot_area()
        .set_format(ChartGradientFill::new().set_gradient_stops(&[
            ChartGradientStop::new("#FFEFD1", 0),
            ChartGradientStop::new("#F0EBD5", 50),
            ChartGradientStop::new("#B69F66", 100),
        ]));

    // Add some axis labels.
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Turn off the chart legend.
    chart.legend().set_hidden();

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(0, 3, &chart, 25, 10)?;

    workbook.save("chart_gradient.xlsx")?;

    Ok(())
}
```

[`ChartFormat`]: crate::ChartFormat
[`ChartGradientFill`]: crate::ChartGradientFill


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


# Chart: Chart data table

An example of creating Excel Column charts with data tables using the
`rust_xlsxwriter` library.


**Image of the output file:**

Chart 1 in the following code is a Column chart with a default chart data table.
<img src="https://rustxlsxwriter.github.io/images/chart_data_table1.png">

Chart 2 in the following code is a Column chart with a chart data table with legend keys.
<img src="https://rustxlsxwriter.github.io/images/chart_data_table2.png">



**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_data_table.rs


use rust_xlsxwriter::{Chart, ChartDataTable, ChartType, Format, Workbook, XlsxError};

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
    // Create a column chart with a data table.
    // -----------------------------------------------------------------------

    // Create a new Column chart.
    let mut chart = Chart::new(ChartType::Column);

    // Configure some data series.
    chart
        .add_series()
        .set_name("Sheet1!$B$1")
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7");

    chart
        .add_series()
        .set_name("Sheet1!$C$1")
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Add a chart title and some axis labels.
    chart.title().set_name("Chart with Data Table");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set a default data table on the X-Axis.
    let table = ChartDataTable::default();
    chart.set_data_table(&table);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a column chart with a data table and legend keys.
    // -----------------------------------------------------------------------

    // Create a new Column chart.
    let mut chart = Chart::new(ChartType::Column);

    // Configure some data series.
    chart
        .add_series()
        .set_name("Sheet1!$B$1")
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7");

    chart
        .add_series()
        .set_name("Sheet1!$C$1")
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Add a chart title and some axis labels.
    chart.title().set_name("Data Table with legend keys");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set a data table on the X-Axis with the legend keys shown.
    let table = ChartDataTable::new().show_legend_keys(true);
    chart.set_data_table(&table);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(17, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Save and close the file.
    // -----------------------------------------------------------------------
    workbook.save("chart_data_table.xlsx")?;

    Ok(())
}
```


# Chart: Chart data tools

A demo of the various Excel chart data tools that are available via the
`rust_xlsxwriter` library.


**Image of the output file:**

Chart 1 in the following code is a trendline chart:
<img src="https://rustxlsxwriter.github.io/images/chart_data_tools1.png">

Chart 2 in the following code is an example of a chart with data labels and markers:
<img src="https://rustxlsxwriter.github.io/images/chart_data_tools2.png">

Chart 3 in the following code is an example of a chart with error bars:
<img src="https://rustxlsxwriter.github.io/images/chart_data_tools3.png">

Chart 4 in the following code is an example of a chart with up-down bars:
<img src="https://rustxlsxwriter.github.io/images/chart_data_tools4.png">

Chart 5 in the following code is an example of a chart with high-low lines:
<img src="https://rustxlsxwriter.github.io/images/chart_data_tools5.png">

Chart 6 in the following code is an example of a chart with drop lines:
<img src="https://rustxlsxwriter.github.io/images/chart_data_tools6.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_data_tools.rs

use rust_xlsxwriter::{
    Chart, ChartDataLabel, ChartErrorBars, ChartErrorBarsType, ChartMarker, ChartSolidFill,
    ChartTrendline, ChartTrendlineType, ChartType, Format, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Number", &bold)?;
    worksheet.write_with_format(0, 1, "Data 1", &bold)?;
    worksheet.write_with_format(0, 2, "Data 2", &bold)?;

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
    // Trendline example
    // -----------------------------------------------------------------------

    // Create a new Line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure the first series with a polynomial trendline.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_trendline(ChartTrendline::new().set_type(ChartTrendlineType::Polynomial(3)));

    // Configure the second series with a linear trendline.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_trendline(ChartTrendline::new().set_type(ChartTrendlineType::Linear));

    // Add a chart title.
    chart.title().set_name("Chart with Trendlines");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Data Labels and Markers example.
    // -----------------------------------------------------------------------

    // Create a new Line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure the first series with data labels and markers.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_data_label(ChartDataLabel::new().show_value())
        .set_marker(ChartMarker::new().set_automatic());

    // Configure the second series as default.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Add a chart title.
    chart.title().set_name("Chart with Data Labels and Markers");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(17, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Error Bar example.
    // -----------------------------------------------------------------------

    // Create a new Line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure the first series with error bars.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_y_error_bars(ChartErrorBars::new().set_type(ChartErrorBarsType::StandardError));

    // Configure the second series as default.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Add a chart title.
    chart.title().set_name("Chart with Error Bars");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(33, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Up-Down Bar example.
    // -----------------------------------------------------------------------

    // Create a new Line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Add the chart up-down bars.
    chart
        .set_up_down_bars(true)
        .set_up_bar_format(ChartSolidFill::new().set_color("#00B050"))
        .set_down_bar_format(ChartSolidFill::new().set_color("#FF0000"));

    // Add a chart title.
    chart.title().set_name("Chart with Up-Down Bars");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(49, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // High-Low Lines example.
    // -----------------------------------------------------------------------

    // Create a new Line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Add the chart High-Low lines.
    chart.set_high_low_lines(true);

    // Add a chart title.
    chart.title().set_name("Chart with High-Low Lines");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(65, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Drop Lines example.
    // -----------------------------------------------------------------------

    // Create a new Line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Add the chart Drop lines.
    chart.set_drop_lines(true);

    // Add a chart title.
    chart.title().set_name("Chart with Drop Lines");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(81, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Save and close the file.
    // -----------------------------------------------------------------------
    workbook.save("chart_data_tools.xlsx")?;

    Ok(())
}
```


# Chart: Clustered categories

Example of creating a clustered Excel chart where there are two levels of
category on the X axis.

The categories in clustered charts are 2D ranges, instead of the more normal
1D ranges.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_chart_clustered.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_clustered.rs

use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Types", &bold)?;
    worksheet.write_with_format(0, 1, "Sub Type", &bold)?;
    worksheet.write_with_format(0, 2, "Value 1", &bold)?;
    worksheet.write_with_format(0, 3, "Value 2", &bold)?;
    worksheet.write_with_format(0, 4, "Value 3", &bold)?;

    worksheet.write(1, 0, "Type 1")?;
    worksheet.write(1, 1, "Sub Type A")?;
    worksheet.write(2, 1, "Sub Type B")?;
    worksheet.write(3, 1, "Sub Type C")?;

    worksheet.write(4, 0, "Type 2")?;
    worksheet.write(4, 1, "Sub Type D")?;
    worksheet.write(5, 1, "Sub Type E")?;

    worksheet.write(1, 2, 5000)?;
    worksheet.write(2, 2, 2000)?;
    worksheet.write(3, 2, 250)?;
    worksheet.write(4, 2, 6000)?;
    worksheet.write(5, 2, 500)?;

    worksheet.write(1, 3, 8000)?;
    worksheet.write(2, 3, 3000)?;
    worksheet.write(3, 3, 1000)?;
    worksheet.write(4, 4, 6500)?;
    worksheet.write(5, 3, 300)?;

    worksheet.write(1, 4, 6000)?;
    worksheet.write(2, 4, 4000)?;
    worksheet.write(3, 4, 2000)?;
    worksheet.write(4, 3, 6000)?;
    worksheet.write(5, 4, 200)?;

    // Create a new chart object.
    let mut chart = Chart::new(ChartType::Column);

    // Configure the series. Note, that the categories are 2D ranges (from
    // column A to column B). This creates the clusters.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$B$6")
        .set_values("Sheet1!$C$2:$C$6");

    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$B$6")
        .set_values("Sheet1!$D$2:$D$6");

    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$B$6")
        .set_values("Sheet1!$E$2:$E$6");

    // Set the Excel chart style.
    chart.set_style(37);

    // Turn off the legend.
    chart.legend().set_hidden();

    // Insert the chart into the worksheet.
    worksheet.insert_chart(2, 6, &chart)?;

    workbook.save("chart_clustered.xlsx")?;

    Ok(())
}
```



# Chart: Gauge Chart


A Gauge Chart isn't a native chart type in Excel. It is constructed by combining
a doughnut chart and a pie chart and by using some non-filled elements to hide
parts of the default charts. This example follows the following online example
of how to create a [Gauge Chart] in Excel.

[Gauge Chart]: https://www.excel-easy.com/examples/gauge-chart.html

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_chart_gauge.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_chart_gauge.rs

use rust_xlsxwriter::{
    Chart, ChartFormat, ChartPoint, ChartSolidFill, ChartType, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    // Add some data for the Doughnut and Pie charts. This is set up so the
    // gauge goes from 0-100. It is initially set at 75%.
    worksheet.write(1, 7, "Donut")?;
    worksheet.write(1, 8, "Pie")?;
    worksheet.write_column(2, 7, [25, 50, 25, 100])?;
    worksheet.write_column(2, 8, [75, 1, 124])?;

    // Configure the doughnut chart as the background for the gauge. We add some
    // custom colors for the Red-Orange-Green of the dial and one non-filled segment.
    let mut chart_doughnut = Chart::new(ChartType::Doughnut);

    let points = vec![
        ChartPoint::new().set_format(
            ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FF0000")),
        ),
        ChartPoint::new().set_format(
            ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FFC000")),
        ),
        ChartPoint::new().set_format(
            ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#00B050")),
        ),
        ChartPoint::new().set_format(ChartFormat::new().set_no_fill()),
    ];

    // Add the chart series.
    chart_doughnut
        .add_series()
        .set_values(("Sheet1", 2, 7, 5, 7))
        .set_name(("Sheet1", 1, 7))
        .set_points(&points);

    // Turn off the chart legend.
    chart_doughnut.legend().set_hidden();

    // Rotate chart so the gauge parts are above the horizontal.
    chart_doughnut.set_rotation(270);

    // Turn off the chart fill and border.
    chart_doughnut
        .chart_area()
        .set_format(ChartFormat::new().set_no_fill().set_no_border());

    // Configure a pie chart as the needle for the gauge.
    let mut chart_pie = Chart::new(ChartType::Pie);
    let points = vec![
        ChartPoint::new().set_format(ChartFormat::new().set_no_fill()),
        ChartPoint::new().set_format(
            ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FF0000")),
        ),
        ChartPoint::new().set_format(ChartFormat::new().set_no_fill()),
    ];

    // Add the chart series.
    chart_pie
        .add_series()
        .set_values(("Sheet1", 2, 8, 5, 8))
        .set_name(("Sheet1", 1, 8))
        .set_points(&points);

    // Rotate the pie chart/needle to align with the doughnut/gauge.
    chart_pie.set_rotation(270);

    // Combine the pie and doughnut charts.
    chart_doughnut.combine(&chart_pie);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(0, 0, &chart_doughnut)?;

    workbook.save("chart_gauge.xlsx")?;

    Ok(())
}
```



# Chart: Chartsheet

In Excel a chartsheet is a type of worksheet that it used primarily to display
one chart. It also supports some other worksheet display options such as headers
and footers, margins, tab selection and print properties

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_chartsheet.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_chartsheet.rs

use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the chart will refer to.
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

    // Create a new bar chart.
    let mut chart = Chart::new(ChartType::Bar);

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title and some axis labels.
    chart.title().set_name("Results of sample analysis");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set an Excel chart style.
    chart.set_style(11);

    // Create a chartsheet.
    let chartsheet = workbook.add_chartsheet();

    // Add the chart to the chartsheet. The row/col position is ignored.
    chartsheet.insert_chart(0, 0, &chart)?;

    // Make the chartsheet the first sheet visible in the workbook.
    chartsheet.set_active(true);

    workbook.save("chartsheet.xlsx")?;

    Ok(())
}
```



# Grouped Rows: Create a grouped row outline

An example of how to group rows into outlines with the `rust_xlsxwriter`
library.

In Excel an outline is a group of rows or columns that can be collapsed or
expanded to simplify hierarchical data. It is often used with the `SUBTOTAL()`
function.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_grouped_rows.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_grouped_rows.rs

use rust_xlsxwriter::{Format, Workbook, Worksheet, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a format to use in headings.
    let bold = Format::new().set_bold();

    // -----------------------------------------------------------------------
    // 1. Add an outline row group with sub-total.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Simple outline row grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add grouping for the over the sub-total range.
    worksheet.group_rows(1, 10)?;

    // -----------------------------------------------------------------------
    // 2. Add nested outline row groups with sub-totals.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Nested outline row grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add grouping for the over the sub-total range.
    worksheet.group_rows(1, 10)?;

    // Add secondary groups within the first range.
    worksheet.group_rows(1, 4)?;
    worksheet.group_rows(6, 9)?;

    // -----------------------------------------------------------------------
    // 3. Add a collapsed inner outline row groups.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Collapsed inner row grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add grouping for the over the sub-total range.
    worksheet.group_rows(1, 10)?;

    // Add collapsed secondary groups within the first range.
    worksheet.group_rows_collapsed(1, 4)?;
    worksheet.group_rows_collapsed(6, 9)?;

    // -----------------------------------------------------------------------
    // 4. Add a collapsed outer row group.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Collapsed outer row grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add collapsed grouping for the over the sub-total range.
    worksheet.group_rows_collapsed(1, 10)?;

    // Add secondary groups within the first range.
    worksheet.group_rows(1, 4)?;
    worksheet.group_rows(6, 9)?;

    // -----------------------------------------------------------------------
    // 5. Row groups with outline symbols on top.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Outline row grouping symbols on top.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add outline row groups.
    worksheet.group_rows(1, 4)?;
    worksheet.group_rows(6, 9)?;

    // Change the worksheet group setting so the outline symbols are on top.
    worksheet.group_symbols_above(true);

    // -----------------------------------------------------------------------
    // 6. Demonstrate all group levels.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Excel outline levels.";
    let levels = [
        "Level 1", "Level 2", "Level 3", "Level 4", //
        "Level 5", "Level 6", "Level 7", "Level 6", //
        "Level 5", "Level 4", "Level 3", "Level 2", //
        "Level 1",
    ];
    worksheet.write_column(0, 0, levels)?;
    worksheet.write_with_format(0, 3, description, &bold)?;

    // Add outline row groups from outer to inner.
    worksheet.group_rows(0, 12)?;
    worksheet.group_rows(1, 11)?;
    worksheet.group_rows(2, 10)?;
    worksheet.group_rows(3, 9)?;
    worksheet.group_rows(4, 8)?;
    worksheet.group_rows(5, 7)?;
    worksheet.group_rows(6, 6)?;

    // Save the file to disk.
    workbook.save("grouped_rows.xlsx")?;

    Ok(())
}

// Generate worksheet data.
pub fn populate_worksheet_data(
    worksheet: &mut Worksheet,
    description: &str,
    bold: &Format,
) -> Result<(), XlsxError> {
    worksheet.write_with_format(0, 3, description, bold)?;

    worksheet.write_with_format(0, 0, "Region", bold)?;
    worksheet.write(1, 0, "North 1")?;
    worksheet.write(2, 0, "North 2")?;
    worksheet.write(3, 0, "North 3")?;
    worksheet.write(4, 0, "North 4")?;
    worksheet.write_with_format(5, 0, "North Total", bold)?;

    worksheet.write_with_format(0, 1, "Sales", bold)?;
    worksheet.write(1, 1, 1000)?;
    worksheet.write(2, 1, 1200)?;
    worksheet.write(3, 1, 900)?;
    worksheet.write(4, 1, 1200)?;
    worksheet.write_formula_with_format(5, 1, "=SUBTOTAL(9,B2:B5)", bold)?;

    worksheet.write(6, 0, "South 1")?;
    worksheet.write(7, 0, "South 2")?;
    worksheet.write(8, 0, "South 3")?;
    worksheet.write(9, 0, "South 4")?;
    worksheet.write_with_format(10, 0, "South Total", bold)?;

    worksheet.write(6, 1, 400)?;
    worksheet.write(7, 1, 600)?;
    worksheet.write(8, 1, 500)?;
    worksheet.write(9, 1, 600)?;
    worksheet.write_formula_with_format(10, 1, "=SUBTOTAL(9,B7:B10)", bold)?;

    worksheet.write_with_format(11, 0, "Grand Total", bold)?;
    worksheet.write_formula_with_format(11, 1, "=SUBTOTAL(9,B2:B11)", bold)?;

    // Autofit the columns for clarity.
    worksheet.autofit();

    Ok(())
}
```


# Grouped Columns: Create a grouped column outline

An example of how to group columns into outlines with the `rust_xlsxwriter`
library.

In Excel an outline is a group of rows or columns that can be collapsed or
expanded to simplify hierarchical data. It is often used with the `SUBTOTAL()`
function.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_grouped_columns.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_grouped_columns.rs

use rust_xlsxwriter::{Format, Workbook, Worksheet, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a format to use in headings.
    let bold = Format::new().set_bold();

    // -----------------------------------------------------------------------
    // 1. Add an outline column group with sub-total.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Simple outline column grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add grouping for the over the sub-total range.
    worksheet.group_columns(1, 8)?;

    // -----------------------------------------------------------------------
    // 2. Add nested outline column groups with sub-totals.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Nested outline column grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add grouping for the over the sub-total range.
    worksheet.group_columns(1, 8)?;

    // Add secondary groups within the first range.
    worksheet.group_columns(1, 3)?;
    worksheet.group_columns(5, 7)?;

    // -----------------------------------------------------------------------
    // 3. Add a collapsed inner outline column groups.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Collapsed inner column grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add grouping for the over the sub-total range.
    worksheet.group_columns(1, 8)?;

    // Add collapsed secondary groups within the first range.
    worksheet.group_columns_collapsed(1, 3)?;
    worksheet.group_columns_collapsed(5, 7)?;

    // -----------------------------------------------------------------------
    // 4. Add a collapsed outer column group.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Collapsed outer column grouping.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add collapsed grouping for the over the sub-total range.
    worksheet.group_columns_collapsed(1, 8)?;

    // Add secondary groups within the first range.
    worksheet.group_columns(1, 3)?;
    worksheet.group_columns(5, 7)?;

    // -----------------------------------------------------------------------
    // 5. Column groups with outline symbols on top.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Outline column grouping symbols to the left.";
    populate_worksheet_data(worksheet, description, &bold)?;

    // Add outline column groups.
    worksheet.group_columns(1, 3)?;
    worksheet.group_columns(5, 7)?;

    // Change the worksheet group setting so outline symbols are to the left.
    worksheet.group_symbols_to_left(true);

    // -----------------------------------------------------------------------
    // 6. Demonstrate all group levels.
    // -----------------------------------------------------------------------

    // Add a worksheet with some sample data.
    let worksheet = workbook.add_worksheet();
    let description = "Excel outline levels.";
    let levels = [
        "Level 1", "Level 2", "Level 3", "Level 4", //
        "Level 5", "Level 6", "Level 7", "Level 6", //
        "Level 5", "Level 4", "Level 3", "Level 2", //
        "Level 1",
    ];
    worksheet.write_row(0, 0, levels)?;
    worksheet.write_with_format(2, 0, description, &bold)?;

    // Add outline column groups from outer to inner.
    worksheet.group_columns(0, 12)?;
    worksheet.group_columns(1, 11)?;
    worksheet.group_columns(2, 10)?;
    worksheet.group_columns(3, 9)?;
    worksheet.group_columns(4, 8)?;
    worksheet.group_columns(5, 7)?;
    worksheet.group_columns(6, 6)?;

    // Save the file to disk.
    workbook.save("grouped_columns.xlsx")?;

    Ok(())
}

// Generate worksheet data.
pub fn populate_worksheet_data(
    worksheet: &mut Worksheet,
    description: &str,
    bold: &Format,
) -> Result<(), XlsxError> {
    worksheet.write_with_format(0, 0, "Region", bold)?;
    worksheet.write_with_format(1, 0, "North", bold)?;
    worksheet.write_with_format(2, 0, "South", bold)?;
    worksheet.write_with_format(3, 0, "East", bold)?;
    worksheet.write_with_format(4, 0, "West", bold)?;

    worksheet.write_with_format(0, 1, "Jan", bold)?;
    worksheet.write(1, 1, 50)?;
    worksheet.write(2, 1, 10)?;
    worksheet.write(3, 1, 45)?;
    worksheet.write(4, 1, 15)?;

    worksheet.write_with_format(0, 2, "Feb", bold)?;
    worksheet.write(1, 2, 20)?;
    worksheet.write(2, 2, 20)?;
    worksheet.write(3, 2, 75)?;
    worksheet.write(4, 2, 15)?;

    worksheet.write_with_format(0, 3, "Mar", bold)?;
    worksheet.write(1, 3, 15)?;
    worksheet.write(2, 3, 30)?;
    worksheet.write(3, 3, 50)?;
    worksheet.write(4, 3, 35)?;

    worksheet.write_with_format(0, 4, "Q1 Total", bold)?;
    worksheet.write_formula_with_format(1, 4, "=SUBTOTAL(9,B2:D2)", bold)?;
    worksheet.write_formula_with_format(2, 4, "=SUBTOTAL(9,B3:D3)", bold)?;
    worksheet.write_formula_with_format(3, 4, "=SUBTOTAL(9,B4:D4)", bold)?;
    worksheet.write_formula_with_format(4, 4, "=SUBTOTAL(9,B5:D5)", bold)?;

    worksheet.write_with_format(0, 5, "Apr", bold)?;
    worksheet.write(1, 5, 25)?;
    worksheet.write(2, 5, 50)?;
    worksheet.write(3, 5, 15)?;
    worksheet.write(4, 5, 35)?;

    worksheet.write_with_format(0, 6, "May", bold)?;
    worksheet.write(1, 6, 65)?;
    worksheet.write(2, 6, 50)?;
    worksheet.write(3, 6, 75)?;
    worksheet.write(4, 6, 70)?;

    worksheet.write_with_format(0, 7, "Jun", bold)?;
    worksheet.write(1, 7, 80)?;
    worksheet.write(2, 7, 50)?;
    worksheet.write(3, 7, 90)?;
    worksheet.write(4, 7, 50)?;

    worksheet.write_with_format(0, 8, "Q2 Total", bold)?;
    worksheet.write_formula_with_format(1, 8, "=SUBTOTAL(9,F2:H2)", bold)?;
    worksheet.write_formula_with_format(2, 8, "=SUBTOTAL(9,F3:H3)", bold)?;
    worksheet.write_formula_with_format(3, 8, "=SUBTOTAL(9,F4:H4)", bold)?;
    worksheet.write_formula_with_format(4, 8, "=SUBTOTAL(9,F5:H5)", bold)?;

    worksheet.write_with_format(0, 9, "H1 Total", bold)?;
    worksheet.write_formula_with_format(1, 9, "=SUBTOTAL(9,B2:I2)", bold)?;
    worksheet.write_formula_with_format(2, 9, "=SUBTOTAL(9,B3:I3)", bold)?;
    worksheet.write_formula_with_format(3, 9, "=SUBTOTAL(9,B4:I4)", bold)?;
    worksheet.write_formula_with_format(4, 9, "=SUBTOTAL(9,B5:I5)", bold)?;

    // Autofit the columns for clarity.
    worksheet.autofit();

    worksheet.write_with_format(6, 0, description, bold)?;

    Ok(())
}
```


# Textbox: Inserting Checkboxes in worksheets

Example of inserting boolean checkboxes into a worksheet.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_checkbox.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_checkbox.rs

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create some formats to use in the worksheet.
    let bold = Format::new().set_bold();
    let light_red = Format::new().set_background_color("FFC7CE");
    let light_green = Format::new().set_background_color("C6EFCE");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set the column width for clarity.
    worksheet.set_column_width(0, 30)?;

    // Write some descriptions.
    worksheet.write_with_format(1, 0, "Some simple checkboxes:", &bold)?;
    worksheet.write_with_format(4, 0, "Some checkboxes with cell formats:", &bold)?;

    // Insert some boolean checkboxes to the worksheet.
    worksheet.insert_checkbox(1, 1, false)?;
    worksheet.insert_checkbox(2, 1, true)?;

    // Insert some checkboxes with cell formats.
    worksheet.insert_checkbox_with_format(4, 1, false, &light_red)?;
    worksheet.insert_checkbox_with_format(5, 1, true, &light_green)?;

    // Save the file to disk.
    workbook.save("checkbox.xlsx")?;

    Ok(())
}
```


# Textbox: Inserting Textboxes in worksheets

Example of inserting a textbox shape into a worksheet.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_textbox.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_textbox.rs

use rust_xlsxwriter::{Shape, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Some text to add to the text box.
    let text = "This is an example of adding a textbox with some text in it";

    // Create a textbox shape and add the text.
    let textbox = Shape::textbox().set_text(text);

    // Insert a textbox in a cell.
    worksheet.insert_shape(1, 1, &textbox)?;

    // Save the file to disk.
    workbook.save("textbox.xlsx")?;

    Ok(())
}
```


# Textbox: Ignore Excel cell errors

An example of turning off worksheet cells errors/warnings.


**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_ignore_errors.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_ignore_errors.rs

use rust_xlsxwriter::{Format, IgnoreError, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a format to use in descriptions.gs
    let bold = Format::new().set_bold();

    // Make the column wider for clarity.
    worksheet.set_column_width(1, 16)?;

    // Write some descriptions for the cells.
    worksheet.write_with_format(1, 1, "Warning:", &bold)?;
    worksheet.write_with_format(2, 1, "Warning turned off:", &bold)?;
    worksheet.write_with_format(4, 1, "Warning:", &bold)?;
    worksheet.write_with_format(5, 1, "Warning turned off:", &bold)?;

    // Write strings that looks like numbers. This will cause an Excel warning.
    worksheet.write_string(1, 2, "123")?;
    worksheet.write_string(2, 2, "123")?;

    // Write a divide by zero formula. This will also cause an Excel warning.
    worksheet.write_formula(4, 2, "=1/0")?;
    worksheet.write_formula(5, 2, "=1/0")?;

    // Turn off some of the warnings:
    worksheet.ignore_error(2, 2, IgnoreError::NumberStoredAsText)?;
    worksheet.ignore_error(5, 2, IgnoreError::FormulaError)?;

    // Save the file to disk.
    workbook.save("ignore_errors.xlsx")?;

    Ok(())
}
```


# Sparklines: simple example

Example of adding sparklines to an Excel spreadsheet using the `rust_xlsxwriter`
library.

Sparklines are small charts that fit in a single cell and are used to show
trends in data. This example shows the basic sparkline types.


**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/sparklines1.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_sparklines1.rs

use rust_xlsxwriter::{Sparkline, SparklineType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Some sample data to plot.
    let data = [[-2, 2, 3, -1, 0], [30, 20, 33, 20, 15], [1, -1, -1, 1, -1]];

    worksheet.write_row_matrix(0, 0, data)?;

    // Add a line sparkline (the default) with markers.
    let sparkline1 = Sparkline::new()
        .set_range(("Sheet1", 0, 0, 0, 4))
        .show_markers(true);

    worksheet.add_sparkline(0, 5, &sparkline1)?;

    // Add a column sparkline with a non-default style.
    let sparkline2 = Sparkline::new()
        .set_range(("Sheet1", 1, 0, 1, 4))
        .set_type(SparklineType::Column)
        .set_style(12);

    worksheet.add_sparkline(1, 5, &sparkline2)?;

    // Add a win/loss sparkline with negative values highlighted.
    let sparkline3 = Sparkline::new()
        .set_range(("Sheet1", 2, 0, 2, 4))
        .set_type(SparklineType::WinLose)
        .show_negative_points(true);

    worksheet.add_sparkline(2, 5, &sparkline3)?;

    // Save the file to disk.
    workbook.save("sparklines1.xlsx")?;

    Ok(())
}
```


# Sparklines: advanced example

Example of adding sparklines to an Excel spreadsheet using the
`rust_xlsxwriter` library.

Sparklines are small charts that fit in a single cell and are used to show
trends in data. This example shows the majority of options that can be applied
to sparklines.


**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/sparklines2.png">


**Code to generate the output file:**

```rust
// Sample code from examples/app_sparklines2.rs

use rust_xlsxwriter::{Format, Sparkline, SparklineType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet1 = workbook.add_worksheet();
    let mut row = 1;

    // Set the columns widths to make the output clearer.
    worksheet1.set_column_width(0, 14)?;
    worksheet1.set_column_width(1, 50)?;
    worksheet1.set_zoom(150);

    // Add some headings.
    let bold = Format::new().set_bold();
    worksheet1.write_with_format(0, 0, "Sparkline", &bold)?;
    worksheet1.write_with_format(0, 1, "Description", &bold)?;

    //
    // Add a default line sparkline.
    //
    let text = "A default line sparkline.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new().set_range(("Sheet2", 0, 0, 0, 9));

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Add a default column sparkline.
    //
    let text = "A default column sparkline.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 1, 0, 1, 9))
        .set_type(SparklineType::Column);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Add a default win/loss sparkline.
    //
    let text = "A default win/loss sparkline.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 2, 0, 2, 9))
        .set_type(SparklineType::WinLose);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 2;

    //
    // Add a line sparkline with markers.
    //
    let text = "Line with markers.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 0, 0, 0, 9))
        .show_markers(true);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Add a line sparkline with high and low points.
    //
    let text = "Line with high and low points.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 0, 0, 0, 9))
        .show_high_point(true)
        .show_low_point(true);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Add a line sparkline with first and last points.
    //
    let text = "Line with first and last point markers.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 0, 0, 0, 9))
        .show_first_point(true)
        .show_last_point(true);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Add a line sparkline with negative point markers.
    //
    let text = "Line with negative point markers.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 0, 0, 0, 9))
        .show_negative_points(true);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Add a line sparkline with axis.
    //
    let text = "Line with axis.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 0, 0, 0, 9))
        .show_axis(true);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 2;

    //
    // Add a column sparkline with style 1. The default style.
    //
    let text = "Column with style 1. The default.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 1, 0, 1, 9))
        .set_type(SparklineType::Column)
        .set_style(1);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Add a column sparkline with style 2.
    //
    let text = "Column with style 2.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 1, 0, 1, 9))
        .set_type(SparklineType::Column)
        .set_style(2);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Add a column sparkline with style 3.
    //
    let text = "Column with style 3.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 1, 0, 1, 9))
        .set_type(SparklineType::Column)
        .set_style(3);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Add a column sparkline with style 4.
    //
    let text = "Column with style 4.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 1, 0, 1, 9))
        .set_type(SparklineType::Column)
        .set_style(4);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Add a column sparkline with style 5.
    //
    let text = "Column with style 5.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 1, 0, 1, 9))
        .set_type(SparklineType::Column)
        .set_style(5);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Add a column sparkline with style 6.
    //
    let text = "Column with style 6.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 1, 0, 1, 9))
        .set_type(SparklineType::Column)
        .set_style(6);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Add a column sparkline with a user defined color.
    //
    let text = "Column with a user defined color.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 1, 0, 1, 9))
        .set_type(SparklineType::Column)
        .set_sparkline_color("#E965E0");

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 2;

    //
    // Add a win/loss sparkline.
    //
    let text = "A win/loss sparkline.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 2, 0, 2, 9))
        .set_type(SparklineType::WinLose);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Add a win/loss sparkline with negative points highlighted.
    //
    let text = "A win/loss sparkline with negative points highlighted.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 2, 0, 2, 9))
        .set_type(SparklineType::WinLose)
        .show_negative_points(true);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 2;

    //
    // Add a left to right (the default) sparkline.
    //
    let text = "A left to right column (the default).";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 3, 0, 3, 9))
        .set_type(SparklineType::Column)
        .set_style(20);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Add a right to left sparkline.
    //
    let text = "A right to left column.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 3, 0, 3, 9))
        .set_type(SparklineType::Column)
        .set_style(20)
        .set_right_to_left(true);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    row += 1;

    //
    // Sparkline and text in one cell. This just requires writing text to the
    // same cell as the sparkline.
    //
    let text = "Sparkline and text in one cell.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 3, 0, 3, 9))
        .set_type(SparklineType::Column)
        .set_style(20);

    worksheet1.add_sparkline(row, 0, &sparkline)?;
    worksheet1.write(row, 0, "Growth")?;
    row += 2;

    //
    // "A grouped sparkline. User changes are applied to all three. Not that the
    // sparkline range is a 2D range and the sparkline is positioned in a 1D
    // range of cells.
    //
    let text = "A grouped sparkline. Changes are applied to all three.";
    worksheet1.write(row, 1, text)?;

    let sparkline = Sparkline::new()
        .set_range(("Sheet2", 4, 0, 6, 9))
        .show_markers(true);

    worksheet1.add_sparkline_group(row, 0, row + 2, 0, &sparkline)?;

    //
    // Add a worksheet with the data to plot on a separate worksheet.
    //
    let worksheet2 = workbook.add_worksheet();

    // Some sample data to plot.
    let data = [
        // Simple line data.
        [-2, 2, 3, -1, 0, -2, 3, 2, 1, 0],
        // Simple column data.
        [30, 20, 33, 20, 15, 5, 5, 15, 10, 15],
        // Simple win/loss data.
        [1, 1, -1, -1, 1, -1, 1, 1, 1, -1],
        // Unbalanced histogram.
        [5, 6, 7, 10, 15, 20, 30, 50, 70, 100],
        // Data for the grouped sparkline example.
        [-2, 2, 3, -1, 0, -2, 3, 2, 1, 0],
        [3, -1, 0, -2, 3, 2, 1, 0, 2, 1],
        [0, -2, 3, 2, 1, 0, 1, 2, 3, 1],
    ];

    worksheet2.write_row_matrix(0, 0, data)?;

    // Save the file to disk.
    workbook.save("sparklines2.xlsx")?;

    Ok(())
}
```


# Traits: Extending generic `write()` to handle user data types

Example of how to extend the the `rust_xlsxwriter`[`Worksheet::write()`] method using the
[`IntoExcelData`] trait to handle arbitrary user data that can be mapped to one
of the main Excel data types.

For this example we create a simple struct type to represent a [Unix Time]. This
is the number of elapsed seconds since the epoch of January 1970 (UTC). Note,
this is for demonstration purposes only. The [`ExcelDateTime`] struct in
 `rust_xlsxwriter` can handle Unix timestamps.


[Unix Time]: https://en.wikipedia.org/wiki/Unix_time
[`IntoExcelData`]: crate::IntoExcelData
[`ExcelDateTime`]: crate::ExcelDateTime
[`Worksheet::write()`]: crate::Worksheet::write

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
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        // Convert the Unix time to an Excel datetime.
        let datetime = 25569.0 + (self.seconds as f64 / (24.0 * 60.0 * 60.0));

        // Write the date with the user supplied format.
        worksheet.write_number_with_format(row, col, datetime, format)
    }
}
```


# Macros: Adding macros to a workbook

An example of adding macros to an `rust_xlsxwriter` file using a VBA macros
file extracted from an existing Excel xlsm file.

The [`vba_extract`](https://crates.io/crates/vba_extract) utility can be used to
extract the `vbaProject.bin` file.

**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_macros.png">

```rust
// Sample code from examples/app_macros.rs

use rust_xlsxwriter::{Button, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add the VBA macro file.
    workbook.add_vba_project("examples/vbaProject.bin")?;

    // Add a worksheet and some text.
    let worksheet = workbook.add_worksheet();

    // Widen the first column for clarity.
    worksheet.set_column_width(0, 30)?;

    worksheet.write(2, 0, "Press the button to say hello:")?;

    // Add a button tied to a macro in the VBA project.
    let button = Button::new()
        .set_caption("Press Me")
        .set_macro("say_hello")
        .set_width(80)
        .set_height(30);

    worksheet.insert_button(2, 1, &button)?;

    // Save the file to disk. Note the `.xlsm` extension. This is required by
    // Excel or it will raise a warning.
    workbook.save("macros.xlsm")?;

    Ok(())
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


# Cell Protection: Setting cell protection in a worksheet

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


# Document Properties: Setting document metadata properties for a workbook

An example of setting workbook document properties for a file created using the
`rust_xlsxwriter` library.

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


# Document Properties: Setting the Sensitivity Label

Sensitivity Labels are a property that can be added to an Office 365 document to
indicate that it is compliant with a company's information protection policies.
Sensitivity Labels have designations like "Confidential", "Internal use only",
or "Public" depending on the policies implemented by the company. They are
generally only enabled for enterprise versions of Office.

See the following Microsoft documentation on how to [Apply sensitivity labels to
your files and email].

Sensitivity Labels are generally stored as custom document properties so they
can be enabled using [`DocProperties::set_custom_property()`]. However, since
the metadata differs from company to company you will need to extract some of
the required metadata from sample files.

[`DocProperties::set_custom_property()`]: crate::DocProperties::set_custom_property

The first step is to create a new file in Excel and set a non-encrypted
sensitivity label. Then unzip the file by changing the extension from `.xlsx` to
`.zip` or by using a command line utility like this:

```bash
$ unzip myfile.xlsx -d myfile
Archive:  myfile.xlsx
  inflating: myfile/[Content_Types].xml
  inflating: myfile/docProps/app.xml
  inflating: myfile/docProps/custom.xml
  inflating: myfile/docProps/core.xml
  inflating: myfile/_rels/.rels
  inflating: myfile/xl/workbook.xml
  inflating: myfile/xl/worksheets/sheet1.xml
  inflating: myfile/xl/styles.xml
  inflating: myfile/xl/theme/theme1.xml
  inflating: myfile/xl/_rels/workbook.xml.rels
```

Then examine the `docProps/custom.xml` file from the unzipped xlsx file. The
file doesn't contain newlines so it is best to view it in an editor that can
handle XML or use a commandline utility like libxml’s [xmllint] to format the
XML for clarity:


```xml
$ xmllint --format myfile/docProps/custom.xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties
    xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
            pid="2"
            name="MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_Enabled">
    <vt:lpwstr>true</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
            pid="3"
            name="MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_SetDate">
    <vt:lpwstr>2024-01-01T12:00:00Z</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
            pid="4"
            name="MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_Method">
    <vt:lpwstr>Privileged</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
            pid="5"
            name="MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_Name">
    <vt:lpwstr>Confidential</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
            pid="6"
            name="MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_SiteId">
    <vt:lpwstr>cb46c030-1825-4e81-a295-151c039dbf02</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
            pid="7"
            name="MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_ActionId">
    <vt:lpwstr>88124cf5-1340-457d-90e1-0000a9427c99</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
            pid="8"
            name="MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_ContentBits">
    <vt:lpwstr>2</vt:lpwstr>
  </property>
</Properties>
```

The MSIP (Microsoft Information Protection) labels in the `name` attributes
contain a GUID that is unique to each company. The `SiteId` field will also be
unique to your company/location. The meaning of each of these fields is
explained in the the following Microsoft document on [Microsoft Information
Protection SDK - Metadata]. Once you have identified the necessary metadata you
can add it to a new document as shown below.

Note, some sensitivity labels require that the document is encrypted. In order
to extract the required metadata you will need to unencrypt the file which may
remove the sensitivity label. In that case you may need to use a third party
tool such as [msoffice-crypt].

[xmllint]: http://xmlsoft.org/xmllint.html

[msoffice-crypt]: https://github.com/herumi/msoffice

[Apply sensitivity labels to your files and email]: https://support.microsoft.com/en-us/office/apply-sensitivity-labels-to-your-files-and-email-2f96e7cd-d5a4-403b-8bd7-4cc636bae0f9

[Microsoft Information Protection SDK - Metadata]: https://learn.microsoft.com/en-us/information-protection/develop/concept-mip-metadata


**Image of the output file:**

<img src="https://rustxlsxwriter.github.io/images/app_sensitivity_label.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_sensitivity_label.rs

use rust_xlsxwriter::{DocProperties, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Metadata extracted from a company specific file.
    let site_id = "cb46c030-1825-4e81-a295-151c039dbf02";
    let action_id = "88124cf5-1340-457d-90e1-0000a9427c99";
    let company_guid = "2096f6a2-d2f7-48be-b329-b73aaa526e5d";

    // Add the document properties. Note that these should all be in text format.
    let properties = DocProperties::new()
        .set_custom_property(format!("MSIP_Label_{company_guid}_Method"), "Privileged")
        .set_custom_property(format!("MSIP_Label_{company_guid}_Name"), "Confidential")
        .set_custom_property(format!("MSIP_Label_{company_guid}_SiteId"), site_id)
        .set_custom_property(format!("MSIP_Label_{company_guid}_ActionId"), action_id)
        .set_custom_property(format!("MSIP_Label_{company_guid}_ContentBits"), "2");

    workbook.set_properties(&properties);

    workbook.save("sensitivity_label.xlsx")?;

    Ok(())
}
```


# Internal links: Creating a Table of Contents

This is an example of creating a "Table of Contents" worksheet with links to
other worksheets in the workbook.

**Image of the output file:**


<img src="https://rustxlsxwriter.github.io/images/app_table_of_contents.png">

**Code to generate the output file:**

```rust
// Sample code from examples/app_table_of_contents.rs

use rust_xlsxwriter::{utility::quote_sheet_name, Format, Url, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create a table of contents worksheet at the start. If the worksheet names
    // are known in advance you can do add them here. For the sake of this
    // example we will assume that they aren't known and/or are created
    // dynamically.
    let _ = workbook.add_worksheet().set_name("Overview")?;

    // Add some worksheets.
    let _ = workbook.add_worksheet().set_name("Pricing")?;
    let _ = workbook.add_worksheet().set_name("Sales")?;
    let _ = workbook.add_worksheet().set_name("Revenue")?;
    let _ = workbook.add_worksheet().set_name("Analytics")?;

    // If the sheet names aren't known in advance we can find them as follows:
    let mut worksheet_names = workbook
        .worksheets()
        .iter()
        .map(|worksheet| worksheet.name())
        .collect::<Vec<_>>();

    // Remove the "Overview" worksheet name.
    worksheet_names.remove(0);

    // Get the "Overview" worksheet to add the table of contents.
    let worksheet = workbook.worksheet_from_name("Overview")?;

    // Write a header.
    let header = Format::new().set_bold().set_background_color("C6EFCE");
    worksheet.write_string_with_format(0, 0, "Table of Contents", &header)?;

    // Write the worksheet names with links.
    for (i, name) in worksheet_names.iter().enumerate() {
        let sheet_name = quote_sheet_name(name);
        let link = format!("internal:{sheet_name}!A1");
        let url = Url::new(link).set_text(name);

        worksheet.write_url(i as u32 + 1, 0, &url)?;
    }

    // Autofit the data for clarity.
    worksheet.autofit();

    // Save the file to disk.
    workbook.save("table_of_contents.xlsx")?;

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
```


# Hyperlinks: Add hyperlinks to a worksheet

This is an example of a program to create demonstrate creating links in a
worksheet using the `rust_xlsxwriter` library.

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

    // Write some URL links.
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


# Excel `LAMBDA()` function: Example of using the Excel 365 `LAMBDA()` function

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


*/
