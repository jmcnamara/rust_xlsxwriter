// RichValue unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod rich_value_tests {

    use crate::rich_value::RichValue;
    use crate::test_functions::xml_to_vec;
    use crate::{xmlwriter, Image};
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble_with_escapes() {
        let image = Image::new("tests/input/images/red.png")
            .unwrap()
            .set_alt_text("a_x0041_b");
        let images = vec![image];

        let mut rich_value = RichValue::new(&images);

        rich_value.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&rich_value.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
              <rv s="0">
                <v>0</v>
                <v>5</v>
                <v>a_x005F_x0041_b</v>
              </rv>
            </rvData>
            "#,
        );

        assert_eq!(expected, got);
    }
}
