# rust_xlsxwriter

The `rust_xlsxwriter` library is a rust library for writing Excel files in
the xlsx format.

<img src="https://rustxlsxwriter.github.io/images/demo.png">

Rust_xlsxwriter is a rust library that can be used to write text, numbers,
dates and formulas to multiple worksheets in a new Excel 2007+ xlsx file. It
has a focus on performance and on fidelity with the file format created by
Excel. It cannot be used to modify an existing file.

## Example

Sample code to generate the Excel file shown above.

```rust
use chrono::NaiveDate;
use rust_xlsxwriter::{Format, Image, Workbook, FormatAlign, FormatBorder, XlsxError};

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
    worksheet.write_string_only(0, 0, "Hello")?;

    // Write a string with the bold format defined above.
    worksheet.write_string(1, 0, "World", &bold_format)?;

    // Write some numbers.
    worksheet.write_number_only(2, 0, 1)?;
    worksheet.write_number_only(3, 0, 2.34)?;

    // Write a number with formatting.
    worksheet.write_number(4, 0, 3.00, &decimal_format)?;

    // Write a formula.
    worksheet.write_formula_only(5, 0, "=SIN(PI()/4)")?;

    // Write a date.
    let date = NaiveDate::from_ymd_opt(2023, 1, 25).unwrap();
    worksheet.write_date(6, 0, &date, &date_format)?;

    // Write some links.
    worksheet.write_url(7, 0, "https://www.rust-lang.org")?;
    worksheet.write_url_with_text(8, 0, "https://www.rust-lang.org", "Learn Rust")?;

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

Rust_xlsxwriter is a port of the [XlsxWriter] Python module by the same
author. Feature porting is a work in progress. The currently supported
features are:

- Support for writing all basic Excel data types.
- Full cell formatting support.
- Formula support, including new Excel 365 dynamic functions.
- Hyperlink support.
- Page/Printing Setup support.
- Merged ranges.
- Worksheet PNG/JPEG/GIF/BMP images.
- Rich multi-format strings.
- Defined names.
- Autofilters.

Rust_xlsxwriter is under active development and new features will be added
frequently.

[XlsxWriter]: https://xlsxwriter.readthedocs.io/index.html
[rust_xlsxwriter GitHub]: https://github.com/jmcnamara/rust_xlsxwriter

## See also

- [User Guide]: Working with the `rust_xlsxwriter` library.
    - [Getting started]: A simple getting started guide on how to use
      `rust_xlsxwriter` in a project and write a Hello World example.
    - [Tutorial]: A larger example of using `rust_xlsxwriter` to write some
       expense data to a spreadsheet.
    - [Cookbook Examples].
- [The rust_xlsxwriter crate]
- [The rust_xlsxwriter API docs at docs.rs]
- [The rust_xlsxwriter repository]
- [Release Notes]

[User Guide]: https://rustxlsxwriter.github.io/index.html
[Getting started]: https://rustxlsxwriter.github.io/getting_stared.html
[Tutorial]: https://rustxlsxwriter.github.io/tutorial/intro.html
[Cookbook Examples]: https://rustxlsxwriter.github.io/examples/intro.html
[The rust_xlsxwriter crate]: https://crates.io/crates/rust_xlsxwriter
[The rust_xlsxwriter API docs at docs.rs]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/
[The rust_xlsxwriter repository]: https://github.com/jmcnamara/rust_xlsxwriter
[Release Notes]: https://rustxlsxwriter.github.io/changelog.html
