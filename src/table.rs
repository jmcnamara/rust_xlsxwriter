// table - A module for creating the Excel Table.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::fmt;

use crate::{
    utility::{self, ToXmlBoolean},
    xmlwriter::XMLWriter,
    ColNum, Formula, RowNum, XlsxError, COL_MAX, ROW_MAX,
};

/// A struct to represent a Table.
///
/// TODO.
#[allow(dead_code)] // TODO
#[derive(Clone)]
pub struct Table {
    pub(crate) writer: XMLWriter,

    pub(crate) columns: Vec<TableColumn>,

    pub(crate) index: u32,
    pub(crate) name: String,
    pub(crate) style: TableStyle,

    pub(crate) first_row: RowNum,
    pub(crate) first_col: ColNum,
    pub(crate) last_row: RowNum,
    pub(crate) last_col: ColNum,

    pub(crate) show_header_row: bool,
    pub(crate) show_total_row: bool,
    pub(crate) show_first_column: bool,
    pub(crate) show_last_column: bool,
    pub(crate) show_banded_rows: bool,
    pub(crate) show_banded_columns: bool,
    pub(crate) show_autofilter: bool,
}

#[allow(dead_code)] // TODO
impl Table {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new Table struct instance.
    #[allow(clippy::new_without_default)]
    pub fn new() -> Table {
        let writer = XMLWriter::new();

        Table {
            writer,
            columns: vec![],
            index: 0,
            name: String::new(),
            style: TableStyle::Medium9,
            first_row: ROW_MAX,
            first_col: COL_MAX,
            last_row: 0,
            last_col: 0,
            show_first_column: false,
            show_last_column: false,
            show_banded_rows: true,
            show_banded_columns: false,
            show_autofilter: true,
            show_header_row: true,
            show_total_row: false,
        }
    }

    /// TODO
    pub fn set_total_row(&mut self, enable: bool) -> &mut Table {
        self.show_total_row = enable;
        self
    }

    /// TODO
    pub fn set_header_row(&mut self, enable: bool) -> &mut Table {
        self.show_header_row = enable;
        self
    }

    /// TODO
    pub fn set_banded_rows(&mut self, enable: bool) -> &mut Table {
        self.show_banded_rows = enable;
        self
    }

    /// TODO
    pub fn set_banded_columns(&mut self, enable: bool) -> &mut Table {
        self.show_banded_columns = enable;
        self
    }

    /// TODO
    pub fn set_first_column(&mut self, enable: bool) -> &mut Table {
        self.show_first_column = enable;
        self
    }

    /// TODO
    pub fn set_last_column(&mut self, enable: bool) -> &mut Table {
        self.show_last_column = enable;
        self
    }

    /// TODO
    pub fn set_autofilter(&mut self, enable: bool) -> &mut Table {
        self.show_autofilter = enable;
        self
    }

    /// TODO
    pub fn set_columns(&mut self, columns: &[TableColumn]) -> &mut Table {
        self.columns = columns.to_vec();
        self
    }

    /// TODO
    pub fn set_name(&mut self, name: impl Into<String>) -> &mut Table {
        self.name = name.into();
        self
    }

    /// TODO
    pub fn set_style(&mut self, style: TableStyle) -> &mut Table {
        self.style = style;
        self
    }

    // Truncate or extend (with defaults) the table columns.
    pub(crate) fn initialize_columns(&mut self) -> Result<(), XlsxError> {
        let num_columns = self.last_col - self.first_col + 1;

        self.columns
            .resize_with(num_columns as usize, TableColumn::default);

        for (index, column) in self.columns.iter_mut().enumerate() {
            if column.name.is_empty() {
                column.name = format!("Column{}", index + 1);
            }
        }

        // TODO add check for duplicate column names.
        if self.columns[0].name == "TODO" {
            return Err(XlsxError::TableError(String::from("Todo")));
        }

        Ok(())
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the table element.
        self.write_table();

        if self.show_autofilter && self.show_header_row {
            // Write the autoFilter element.
            self.write_auto_filter();
        }

        // Write the tableColumns element.
        self.write_columns();

        // Write the tableStyleInfo element.
        self.write_table_style_info();

        // Close the table tag.
        self.writer.xml_end_tag("table");
    }

    // Write the <table> element.
    fn write_table(&mut self) {
        let schema = "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string();
        let range =
            utility::cell_range(self.first_row, self.first_col, self.last_row, self.last_col);
        let name = if self.name.is_empty() {
            format!("Table{}", self.index)
        } else {
            self.name.clone()
        };

        let mut attributes = vec![
            ("xmlns", schema),
            ("id", self.index.to_string()),
            ("name", name.clone()),
            ("displayName", name),
            ("ref", range),
        ];

        if !self.show_header_row {
            attributes.push(("headerRowCount", "0".to_string()));
        }

        if self.show_total_row {
            attributes.push(("totalsRowCount", "1".to_string()));
        } else {
            attributes.push(("totalsRowShown", "0".to_string()));
        }

        self.writer.xml_start_tag("table", &attributes);
    }

    // Write the <autoFilter> element.
    fn write_auto_filter(&mut self) {
        let mut last_row = self.last_row;
        if self.show_total_row {
            last_row -= 1;
        }

        let attributes = vec![(
            "ref",
            utility::cell_range(self.first_row, self.first_col, last_row, self.last_col),
        )];

        self.writer.xml_empty_tag("autoFilter", &attributes);
    }

    // Write the <tableColumns> element.
    fn write_columns(&mut self) {
        let attributes = vec![("count", self.columns.len().to_string())];

        self.writer.xml_start_tag("tableColumns", &attributes);

        for (index, column) in self.columns.clone().iter().enumerate() {
            // Write the tableColumn element.
            self.write_column(index + 1, column);
        }

        self.writer.xml_end_tag("tableColumns");
    }

    // Write the <tableColumn> element.
    fn write_column(&mut self, index: usize, column: &TableColumn) {
        let mut attributes = vec![("id", index.to_string()), ("name", column.name.clone())];

        if !column.total_label.is_empty() {
            attributes.push(("totalsRowLabel", column.total_label.clone()));
        } else if column.total_function != TableFunction::None {
            attributes.push(("totalsRowFunction", column.total_function.to_string()));
        }

        self.writer.xml_empty_tag("tableColumn", &attributes);
    }

    // Write the <tableStyleInfo> element.
    fn write_table_style_info(&mut self) {
        let mut attributes = vec![];

        if self.style != TableStyle::None {
            attributes.push(("name", self.style.to_string()));
        }

        attributes.push(("showFirstColumn", self.show_first_column.to_xml_bool()));
        attributes.push(("showLastColumn", self.show_last_column.to_xml_bool()));
        attributes.push(("showRowStripes", self.show_banded_rows.to_xml_bool()));
        attributes.push(("showColumnStripes", self.show_banded_columns.to_xml_bool()));

        self.writer.xml_empty_tag("tableStyleInfo", &attributes);
    }
}

#[allow(dead_code)] // TODO
#[derive(Clone)]
/// TODO
pub struct TableColumn {
    pub(crate) name: String,
    pub(crate) total_function: TableFunction,
    pub(crate) total_label: String,
}

#[allow(dead_code)] // TODO
impl TableColumn {
    /// TODO
    pub fn new() -> TableColumn {
        TableColumn {
            name: String::new(),
            total_function: TableFunction::None,
            total_label: String::new(),
        }
    }

    /// TODO
    pub fn set_header(mut self, name: impl Into<String>) -> TableColumn {
        self.name = name.into();
        self
    }

    /// TODO
    pub fn set_total_function(mut self, function: TableFunction) -> TableColumn {
        self.total_function = function;
        self
    }

    /// TODO
    pub fn set_total_label(mut self, label: impl Into<String>) -> TableColumn {
        self.total_label = label.into();
        self
    }

    // Convert the SUBTOTAL type to a worksheet formula.
    pub(crate) fn total_function(&self) -> Formula {
        let column_name = self
            .name
            .replace('\'', "''")
            .replace('#', "'#")
            .replace(']', "']")
            .replace('[', "'[");

        match self.total_function {
            TableFunction::Max => Formula::new(format!("SUBTOTAL(104,[{column_name}])")),
            TableFunction::Min => Formula::new(format!("SUBTOTAL(105,[{column_name}])")),
            TableFunction::Sum => Formula::new(format!("SUBTOTAL(109,[{column_name}])")),
            TableFunction::Var => Formula::new(format!("SUBTOTAL(110,[{column_name}])")),
            TableFunction::None => Formula::new(""),
            TableFunction::Count => Formula::new(format!("SUBTOTAL(103,[{column_name}])")),
            TableFunction::StdDev => Formula::new(format!("SUBTOTAL(107,[{column_name}])")),
            TableFunction::Average => Formula::new(format!("SUBTOTAL(101,[{column_name}])")),
            TableFunction::CountNumbers => Formula::new(format!("SUBTOTAL(102,[{column_name}])")),
        }
    }
}

impl Default for TableColumn {
    fn default() -> Self {
        Self::new()
    }
}

/// Standard Excel functions for totals in tables.
///
/// Definitions for the standard Excel functions that are available via the
/// dropdown in the total row of an Excel table.
///
/// TODO
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum TableFunction {
    /// The "total row" option is enable but there is no total function.
    None,

    /// Use the average function as the table total.
    Average,

    /// Use the count function as the table total.
    Count,

    /// Use the count numbers function as the table total.
    CountNumbers,

    /// Use the max function as the table total.
    Max,

    /// Use the min function as the table total.
    Min,

    /// Use the standard deviation function as the table total.
    StdDev,

    /// Use the sum function as the table total.
    Sum,

    /// Use the var function as the table total.
    Var,
}

impl fmt::Display for TableFunction {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            TableFunction::Max => write!(f, "max"),
            TableFunction::Min => write!(f, "min"),
            TableFunction::Sum => write!(f, "sum"),
            TableFunction::Var => write!(f, "var"),
            TableFunction::None => write!(f, "None"),
            TableFunction::Count => write!(f, "count"),
            TableFunction::StdDev => write!(f, "stdDev"),
            TableFunction::Average => write!(f, "average"),
            TableFunction::CountNumbers => write!(f, "countNums"),
        }
    }
}

/// TODO
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum TableStyle {
    /// No table style.
    None,

    /// Table Style Light 1, White.
    Light1,

    /// Table Style Light 2, Light Blue.
    Light2,

    /// Table Style Light 3, Light Orange.
    Light3,

    /// Table Style Light 4, White.
    Light4,

    /// Table Style Light 5, Light Yellow.
    Light5,

    /// Table Style Light 6, Light Blue.
    Light6,

    /// Table Style Light 7, Light Green.
    Light7,

    /// Table Style Light 8, White.
    Light8,

    /// Table Style Light 9, Blue.
    Light9,

    /// Table Style Light 10, Orange.
    Light10,

    /// Table Style Light 11, White.
    Light11,

    /// Table Style Light 12, Gold.
    Light12,

    /// Table Style Light 13, Blue.
    Light13,

    /// Table Style Light 14, Green.
    Light14,

    /// Table Style Light 15, White.
    Light15,

    /// Table Style Light 16, Light Blue.
    Light16,

    /// Table Style Light 17, Light Orange.
    Light17,

    /// Table Style Light 18, White.
    Light18,

    /// Table Style Light 19, Light Yellow.
    Light19,

    /// Table Style Light 20, Light Blue.
    Light20,

    /// Table Style Light 21, Light Green.
    Light21,

    /// Table Style Medium 1, White.
    Medium1,

    /// Table Style Medium 2, Blue.
    Medium2,

    /// Table Style Medium 3, Orange.
    Medium3,

    /// Table Style Medium 4, White.
    Medium4,

    /// Table Style Medium 5, Gold.
    Medium5,

    /// Table Style Medium 6, Blue.
    Medium6,

    /// Table Style Medium 7, Green.
    Medium7,

    /// Table Style Medium 8, Light Grey.
    Medium8,

    /// Table Style Medium 9, Blue.
    Medium9,

    /// Table Style Medium 10, Orange.
    Medium10,

    /// Table Style Medium 11, Light Grey.
    Medium11,

    /// Table Style Medium 12, Gold.
    Medium12,

    /// Table Style Medium 13, Blue.
    Medium13,

    /// Table Style Medium 14, Green.
    Medium14,

    /// Table Style Medium 15, White.
    Medium15,

    /// Table Style Medium 16, Blue.
    Medium16,

    /// Table Style Medium 17, Orange.
    Medium17,

    /// Table Style Medium 18, White.
    Medium18,

    /// Table Style Medium 19, Gold.
    Medium19,

    /// Table Style Medium 20, Blue.
    Medium20,

    /// Table Style Medium 21, Green.
    Medium21,

    /// Table Style Medium 22, Light Grey.
    Medium22,

    /// Table Style Medium 23, Light Blue.
    Medium23,

    /// Table Style Medium 24, Light Orange.
    Medium24,

    /// Table Style Medium 25, Light Grey.
    Medium25,

    /// Table Style Medium 26, Light Yellow.
    Medium26,

    /// Table Style Medium 27, Light Blue.
    Medium27,

    /// Table Style Medium 28, Light Green.
    Medium28,

    /// Table Style Dark 1, Dark Grey.
    Dark1,

    /// Table Style Dark 2, Dark Blue.
    Dark2,

    /// Table Style Dark 3, Brown.
    Dark3,

    /// Table Style Dark 4, Grey.
    Dark4,

    /// Table Style Dark 5, Dark Yellow.
    Dark5,

    /// Table Style Dark 6, Blue.
    Dark6,

    /// Table Style Dark 7, Dark Green.
    Dark7,

    /// Table Style Dark 8, Light Grey.
    Dark8,

    /// Table Style Dark 9, Light Orange.
    Dark9,

    /// Table Style Dark 10, Gold.
    Dark10,

    /// Table Style Dark 11, Green.
    Dark11,
}

impl fmt::Display for TableStyle {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            TableStyle::None => write!(f, "TableStyleNone"),
            TableStyle::Light1 => write!(f, "TableStyleLight1"),
            TableStyle::Light2 => write!(f, "TableStyleLight2"),
            TableStyle::Light3 => write!(f, "TableStyleLight3"),
            TableStyle::Light4 => write!(f, "TableStyleLight4"),
            TableStyle::Light5 => write!(f, "TableStyleLight5"),
            TableStyle::Light6 => write!(f, "TableStyleLight6"),
            TableStyle::Light7 => write!(f, "TableStyleLight7"),
            TableStyle::Light8 => write!(f, "TableStyleLight8"),
            TableStyle::Light9 => write!(f, "TableStyleLight9"),
            TableStyle::Light10 => write!(f, "TableStyleLight10"),
            TableStyle::Light11 => write!(f, "TableStyleLight11"),
            TableStyle::Light12 => write!(f, "TableStyleLight12"),
            TableStyle::Light13 => write!(f, "TableStyleLight13"),
            TableStyle::Light14 => write!(f, "TableStyleLight14"),
            TableStyle::Light15 => write!(f, "TableStyleLight15"),
            TableStyle::Light16 => write!(f, "TableStyleLight16"),
            TableStyle::Light17 => write!(f, "TableStyleLight17"),
            TableStyle::Light18 => write!(f, "TableStyleLight18"),
            TableStyle::Light19 => write!(f, "TableStyleLight19"),
            TableStyle::Light20 => write!(f, "TableStyleLight20"),
            TableStyle::Light21 => write!(f, "TableStyleLight21"),
            TableStyle::Medium1 => write!(f, "TableStyleMedium1"),
            TableStyle::Medium2 => write!(f, "TableStyleMedium2"),
            TableStyle::Medium3 => write!(f, "TableStyleMedium3"),
            TableStyle::Medium4 => write!(f, "TableStyleMedium4"),
            TableStyle::Medium5 => write!(f, "TableStyleMedium5"),
            TableStyle::Medium6 => write!(f, "TableStyleMedium6"),
            TableStyle::Medium7 => write!(f, "TableStyleMedium7"),
            TableStyle::Medium8 => write!(f, "TableStyleMedium8"),
            TableStyle::Medium9 => write!(f, "TableStyleMedium9"),
            TableStyle::Medium10 => write!(f, "TableStyleMedium10"),
            TableStyle::Medium11 => write!(f, "TableStyleMedium11"),
            TableStyle::Medium12 => write!(f, "TableStyleMedium12"),
            TableStyle::Medium13 => write!(f, "TableStyleMedium13"),
            TableStyle::Medium14 => write!(f, "TableStyleMedium14"),
            TableStyle::Medium15 => write!(f, "TableStyleMedium15"),
            TableStyle::Medium16 => write!(f, "TableStyleMedium16"),
            TableStyle::Medium17 => write!(f, "TableStyleMedium17"),
            TableStyle::Medium18 => write!(f, "TableStyleMedium18"),
            TableStyle::Medium19 => write!(f, "TableStyleMedium19"),
            TableStyle::Medium20 => write!(f, "TableStyleMedium20"),
            TableStyle::Medium21 => write!(f, "TableStyleMedium21"),
            TableStyle::Medium22 => write!(f, "TableStyleMedium22"),
            TableStyle::Medium23 => write!(f, "TableStyleMedium23"),
            TableStyle::Medium24 => write!(f, "TableStyleMedium24"),
            TableStyle::Medium25 => write!(f, "TableStyleMedium25"),
            TableStyle::Medium26 => write!(f, "TableStyleMedium26"),
            TableStyle::Medium27 => write!(f, "TableStyleMedium27"),
            TableStyle::Medium28 => write!(f, "TableStyleMedium28"),
            TableStyle::Dark1 => write!(f, "TableStyleDark1"),
            TableStyle::Dark2 => write!(f, "TableStyleDark2"),
            TableStyle::Dark3 => write!(f, "TableStyleDark3"),
            TableStyle::Dark4 => write!(f, "TableStyleDark4"),
            TableStyle::Dark5 => write!(f, "TableStyleDark5"),
            TableStyle::Dark6 => write!(f, "TableStyleDark6"),
            TableStyle::Dark7 => write!(f, "TableStyleDark7"),
            TableStyle::Dark8 => write!(f, "TableStyleDark8"),
            TableStyle::Dark9 => write!(f, "TableStyleDark9"),
            TableStyle::Dark10 => write!(f, "TableStyleDark10"),
            TableStyle::Dark11 => write!(f, "TableStyleDark11"),
        }
    }
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use crate::table::Table;
    use crate::test_functions::xml_to_vec;
    use crate::{TableColumn, TableFunction};
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble1() {
        let mut table = Table::new();

        table.first_row = 2;
        table.first_col = 2;
        table.last_row = 12;
        table.last_col = 5;
        table.index = 1;

        table.initialize_columns().unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F13" totalsRowShown="0">
                <autoFilter ref="C3:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble2() {
        let mut table = Table::new();

        table.first_row = 3;
        table.first_col = 3;
        table.last_row = 14;
        table.last_col = 8;
        table.index = 2;

        table.set_style(crate::TableStyle::Light17);

        table.initialize_columns().unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="2" name="Table2" displayName="Table2" ref="D4:I15" totalsRowShown="0">
                <autoFilter ref="D4:I15"/>
                <tableColumns count="6">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                    <tableColumn id="5" name="Column5"/>
                    <tableColumn id="6" name="Column6"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleLight17" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble3() {
        let mut table = Table::new();

        table.first_row = 4;
        table.first_col = 2;
        table.last_row = 15;
        table.last_col = 3;
        table.index = 1;

        table.set_first_column(true);
        table.set_last_column(true);
        table.set_banded_rows(false);
        table.set_banded_columns(true);

        table.initialize_columns().unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C5:D16" totalsRowShown="0">
                <autoFilter ref="C5:D16"/>
                <tableColumns count="2">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="1" showLastColumn="1" showRowStripes="0" showColumnStripes="1"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble4() {
        let mut table = Table::new();

        table.first_row = 2;
        table.first_col = 2;
        table.last_row = 12;
        table.last_col = 5;
        table.index = 1;

        table.set_autofilter(false);

        table.initialize_columns().unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F13" totalsRowShown="0">
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble5() {
        let mut table = Table::new();

        table.first_row = 3;
        table.first_col = 2;
        table.last_row = 12;
        table.last_col = 5;
        table.index = 1;

        table.set_header_row(false);

        table.initialize_columns().unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C4:F13" headerRowCount="0" totalsRowShown="0">
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble6() {
        let mut table = Table::new();

        table.first_row = 2;
        table.first_col = 2;
        table.last_row = 12;
        table.last_col = 5;
        table.index = 1;

        let columns = vec![
            TableColumn::new().set_header("Foo"),
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::new().set_header("Baz"),
        ];

        table.set_columns(&columns);

        table.initialize_columns().unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F13" totalsRowShown="0">
                <autoFilter ref="C3:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Foo"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Baz"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);

        // Try a variation with too many columns settings. It should be
        // truncated to the actual number of columns.
        let mut table = Table::new();

        table.first_row = 2;
        table.first_col = 2;
        table.last_row = 12;
        table.last_col = 5;
        table.index = 1;

        let columns = vec![
            TableColumn::new().set_header("Foo"),
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::new().set_header("Baz"),
            TableColumn::new().set_header("Too many"),
        ];

        table.set_columns(&columns);

        table.initialize_columns().unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble7() {
        let mut table = Table::new();

        table.first_row = 2;
        table.first_col = 2;
        table.last_row = 13;
        table.last_col = 5;
        table.index = 1;

        table.set_total_row(true);

        table.initialize_columns().unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F14" totalsRowCount="1">
                <autoFilter ref="C3:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble8() {
        let mut table = Table::new();

        table.first_row = 2;
        table.first_col = 2;
        table.last_row = 13;
        table.last_col = 5;
        table.index = 1;

        let columns = vec![
            TableColumn::new().set_total_label("Total"),
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::new().set_total_function(TableFunction::Count),
        ];

        table.set_columns(&columns);
        table.set_total_row(true);

        table.initialize_columns().unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F14" totalsRowCount="1">
                <autoFilter ref="C3:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1" totalsRowLabel="Total"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4" totalsRowFunction="count"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble9() {
        let mut table = Table::new();

        table.first_row = 1;
        table.first_col = 1;
        table.last_row = 7;
        table.last_col = 10;
        table.index = 1;

        let columns = vec![
            TableColumn::new()
                .set_total_label("Total")
                .set_total_function(TableFunction::Max), // Max should be ignore due to label.
            TableColumn::new().set_total_function(TableFunction::None), // Should be ignored.
            TableColumn::new().set_total_function(TableFunction::Average),
            TableColumn::new().set_total_function(TableFunction::Count),
            TableColumn::new().set_total_function(TableFunction::CountNumbers),
            TableColumn::new().set_total_function(TableFunction::Max),
            TableColumn::new().set_total_function(TableFunction::Min),
            TableColumn::new().set_total_function(TableFunction::Sum),
            TableColumn::new().set_total_function(TableFunction::StdDev),
            TableColumn::new().set_total_function(TableFunction::Var),
        ];

        table.set_columns(&columns);
        table.set_total_row(true);

        table.initialize_columns().unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="B2:K8" totalsRowCount="1">
                <autoFilter ref="B2:K7"/>
                <tableColumns count="10">
                    <tableColumn id="1" name="Column1" totalsRowLabel="Total"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3" totalsRowFunction="average"/>
                    <tableColumn id="4" name="Column4" totalsRowFunction="count"/>
                    <tableColumn id="5" name="Column5" totalsRowFunction="countNums"/>
                    <tableColumn id="6" name="Column6" totalsRowFunction="max"/>
                    <tableColumn id="7" name="Column7" totalsRowFunction="min"/>
                    <tableColumn id="8" name="Column8" totalsRowFunction="sum"/>
                    <tableColumn id="9" name="Column9" totalsRowFunction="stdDev"/>
                    <tableColumn id="10" name="Column10" totalsRowFunction="var"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble10() {
        let mut table = Table::new();

        table.first_row = 1;
        table.first_col = 2;
        table.last_row = 12;
        table.last_col = 5;
        table.index = 1;

        table.set_name("MyTable");

        table.initialize_columns().unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="MyTable" displayName="MyTable" ref="C2:F13" totalsRowShown="0">
                <autoFilter ref="C2:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }
}
