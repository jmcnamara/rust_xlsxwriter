# rust_xlsxwriter_derive

The `rust_xlsxwriter_derive` crate provides the `XlsxSerialize` derived
trait which is used in conjunction with `rust_xlsxwriter` serialization.

`XlsxSerialize` can be used to set container and field attributes for structs to
define Excel formatting and other options when serializing them to Excel using
`rust_xlsxwriter` and `Serde`.

```rust
use rust_xlsxwriter::{Workbook, XlsxError, XlsxSerialize};
use serde::Serialize;

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a serializable struct.
    #[derive(XlsxSerialize, Serialize)]
    #[xlsx(header_format = Format::new().set_bold())]
    struct Produce {
        #[xlsx(rename = "Item", column_width = 12.0)]
        fruit: &'static str,

        #[xlsx(rename = "Price", num_format = "$0.00")]
        cost: f64,
    }

    // Create some data instances.
    let items = [
        Produce {
            fruit: "Peach",
            cost: 1.05,
        },
        Produce {
            fruit: "Plum",
            cost: 0.15,
        },
        Produce {
            fruit: "Pear",
            cost: 0.75,
        },
    ];

    // Set the serialization location and headers.
    worksheet.set_serialize_headers::<Produce>(0, 0)?;

    // Serialize the data.
    worksheet.serialize(&items)?;

    // Save the file to disk.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
```

The output file is shown below. Note the change or column width in Column A,
the renamed headers and the currency format in Column B numbers.

<img src="https://rustxlsxwriter.github.io/images/xlsxserialize_intro.png">

For more information see the documentation on [Working with Serde] in the
`rust_xlsxwriter` docs.


## See also

- [The rust_xlsxwriter crate].
- [The rust_xlsxwriter API docs at docs.rs].
- [The rust_xlsxwriter repository].

[The rust_xlsxwriter crate]: https://crates.io/crates/rust_xlsxwriter
[The rust_xlsxwriter API docs at docs.rs]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/
[The rust_xlsxwriter repository]: https://github.com/jmcnamara/rust_xlsxwriter
[Working with Serde]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/serializer/index.html
