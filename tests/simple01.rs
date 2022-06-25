use rust_xlsxwriter::Workbook;

#[test]
fn create_file() {
    let mut workbook = Workbook::new("tests/output/rs_simple01.xlsx");
    workbook.close();
}
