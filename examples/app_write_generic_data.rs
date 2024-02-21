// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of how to extend the the `rust_xlsxwriter` `write()` method using
//! the IntoExcelData trait to handle arbitrary user data that can be mapped to
//! one of the main Excel data types.

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
