# rust_xlsxwriter

The `rust_xlsxwriter` library is a rust library for writing Excel files in
the XLSX format.

<img src="https://rustxlsxwriter.github.io/images/demo.png">

Rust_xlsxwriter is a rust library that can be used to write text, numbers,
dates and formulas to multiple worksheets in a new Excel 2007+ XLSX file. It
has a focus on performance and on fidelity with the file format created by
Excel. It cannot be used to modify an existing file.

## Example

Sample code to generate the Excel file shown above.

```rust
use chrono::NaiveDate;
use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file.
    let mut workbook = Workbook::new("demo.xlsx");

    // Create some formats to use in the worksheet.
    let bold_format = Format::new().set_bold();
    let decimal_format = Format::new().set_num_format("0.000");
    let date_format = Format::new().set_num_format("yyyy-mm-dd");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set the column width for clarity.
    worksheet.set_column_width(0, 15)?;

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

    // Write the date .
    let date = NaiveDate::from_ymd(2023, 1, 25);
    worksheet.write_date(6, 0, date, &date_format)?;

    workbook.close()?;

    Ok(())
}
```

Rust_xlsxwriter is a port of the [XlsxWriter] Python module by the same
author. Feature porting is a work in progress. The currently supported
features are:

- Support for writing all basic Excel data types.
- Full cell formatting support.

Rust_xlsxwriter is under active development and new features will be added
frequently.

[XlsxWriter]: https://xlsxwriter.readthedocs.io/index.html
[rust_xlsxwriter GitHub]: https://github.com/jmcnamara/rust_xlsxwriter

