use rust_xlsxwriter::Workbook;


#[test]
fn create_file() {
    let mut workbook = Workbook::new("simple01.xlsx");
    workbook.assemble_xml_file();
    workbook.close();
}
