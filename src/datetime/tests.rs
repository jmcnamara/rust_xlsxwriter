// excel_datetime unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod datetime_tests {

    use crate::ExcelDateTime;
    use chrono::prelude::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_timestamp_to_iso8601_times() {
        let tests = [
            ("2000-01-01T01:00:00Z", 946688400),
            ("2000-01-01T00:01:00Z", 946684860),
            ("2000-01-01T00:00:01Z", 946684801),
            ("2000-01-01T23:00:00Z", 946767600),
            ("2000-01-01T00:59:00Z", 946688340),
            ("2000-01-01T00:00:59Z", 946684859),
            ("2000-01-01T23:59:59Z", 946771199),
        ];

        for (expected, unix_time) in tests {
            let got = ExcelDateTime::unix_time_to_iso8601(unix_time);
            assert_eq!(expected, got);
        }
    }

    #[test]
    fn test_timestamp_to_iso8601_dates() {
        let tests = [
            ("1970-01-01T00:00:00Z", 0),
            ("1970-02-28T00:00:00Z", 5011200),
            ("1970-03-01T00:00:00Z", 5097600),
            ("1970-12-01T00:00:00Z", 28857600),
            ("1970-12-31T00:00:00Z", 31449600),
            ("1971-01-01T00:00:00Z", 31536000),
            ("1971-02-28T00:00:00Z", 36547200),
            ("1971-03-01T00:00:00Z", 36633600),
            ("1971-12-01T00:00:00Z", 60393600),
            ("1971-12-31T00:00:00Z", 62985600),
            ("1972-01-01T00:00:00Z", 63072000),
            ("1972-01-02T00:00:00Z", 63158400),
            ("1972-02-26T00:00:00Z", 67910400),
            ("1972-02-27T00:00:00Z", 67996800),
            ("1972-02-28T00:00:00Z", 68083200),
            ("1972-02-29T00:00:00Z", 68169600),
            ("1972-03-01T00:00:00Z", 68256000),
            ("1972-12-01T00:00:00Z", 92016000),
            ("1972-12-30T00:00:00Z", 94521600),
            ("1972-12-31T00:00:00Z", 94608000),
            ("1973-01-01T00:00:00Z", 94694400),
            ("1973-02-28T00:00:00Z", 99705600),
            ("1973-03-01T00:00:00Z", 99792000),
            ("1973-12-01T00:00:00Z", 123552000),
            ("1973-12-31T00:00:00Z", 126144000),
            ("1974-01-01T00:00:00Z", 126230400),
            ("1974-02-28T00:00:00Z", 131241600),
            ("1974-03-01T00:00:00Z", 131328000),
            ("1974-12-01T00:00:00Z", 155088000),
            ("1974-12-31T00:00:00Z", 157680000),
            ("2000-01-01T00:00:00Z", 946684800),
            ("2000-02-28T00:00:00Z", 951696000),
            ("2000-02-29T00:00:00Z", 951782400),
            ("2000-03-01T00:00:00Z", 951868800),
            ("2000-12-01T00:00:00Z", 975628800),
            ("2000-12-31T00:00:00Z", 978220800),
            ("2001-01-01T00:00:00Z", 978307200),
            ("2001-02-28T00:00:00Z", 983318400),
            ("2001-03-01T00:00:00Z", 983404800),
            ("2001-12-01T00:00:00Z", 1007164800),
            ("2001-12-31T00:00:00Z", 1009756800),
            ("2002-01-01T00:00:00Z", 1009843200),
            ("2002-02-28T00:00:00Z", 1014854400),
            ("2002-03-01T00:00:00Z", 1014940800),
            ("2002-12-01T00:00:00Z", 1038700800),
            ("2002-12-31T00:00:00Z", 1041292800),
            ("2003-01-01T00:00:00Z", 1041379200),
            ("2003-02-28T00:00:00Z", 1046390400),
            ("2003-03-01T00:00:00Z", 1046476800),
            ("2003-12-01T00:00:00Z", 1070236800),
            ("2003-12-31T00:00:00Z", 1072828800),
            ("2004-01-01T00:00:00Z", 1072915200),
            ("2004-02-28T00:00:00Z", 1077926400),
            ("2004-02-29T00:00:00Z", 1078012800),
            ("2004-03-01T00:00:00Z", 1078099200),
            ("2004-12-01T00:00:00Z", 1101859200),
            ("2004-12-31T00:00:00Z", 1104451200),
            ("2099-01-01T00:00:00Z", 4070908800),
            ("2099-02-28T00:00:00Z", 4075920000),
            ("2099-03-01T00:00:00Z", 4076006400),
            ("2099-12-01T00:00:00Z", 4099766400),
            ("2099-12-30T00:00:00Z", 4102272000),
            ("2099-12-31T00:00:00Z", 4102358400),
            ("2100-01-01T00:00:00Z", 4102444800),
            ("2100-02-28T00:00:00Z", 4107456000),
            ("2100-03-01T00:00:00Z", 4107542400),
            ("2100-12-01T00:00:00Z", 4131302400),
            ("2100-12-31T00:00:00Z", 4133894400),
            ("2399-01-01T00:00:00Z", 13537929600),
            ("2399-02-28T00:00:00Z", 13542940800),
            ("2399-03-01T00:00:00Z", 13543027200),
            ("2399-12-01T00:00:00Z", 13566787200),
            ("2399-12-31T00:00:00Z", 13569379200),
            ("2400-01-01T00:00:00Z", 13569465600),
            ("2400-02-28T00:00:00Z", 13574476800),
            ("2400-03-01T00:00:00Z", 13574649600),
            ("2400-12-01T00:00:00Z", 13598409600),
            ("2400-12-31T00:00:00Z", 13601001600),
            ("2404-01-01T00:00:00Z", 13695696000),
            ("2404-01-02T00:00:00Z", 13695782400),
            ("2404-01-03T00:00:00Z", 13695868800),
            ("2404-02-28T00:00:00Z", 13700707200),
            ("2404-02-29T00:00:00Z", 13700793600),
            ("2404-03-01T00:00:00Z", 13700880000),
            ("2404-12-01T00:00:00Z", 13724640000),
            ("2404-12-31T00:00:00Z", 13727232000),
            ("4000-01-01T00:00:00Z", 64060588800),
            ("4000-02-28T00:00:00Z", 64065600000),
            ("4000-02-29T00:00:00Z", 64065686400),
            ("4000-03-01T00:00:00Z", 64065772800),
            ("4000-12-01T00:00:00Z", 64089532800),
            ("4000-12-31T00:00:00Z", 64092124800),
            ("9999-01-01T00:00:00Z", 253370764800),
            ("9999-02-28T00:00:00Z", 253375776000),
            ("9999-03-01T00:00:00Z", 253375862400),
            ("9999-12-01T00:00:00Z", 253399622400),
            ("9999-12-31T00:00:00Z", 253402214400),
        ];

        for (expected, unix_time) in tests {
            let got = ExcelDateTime::unix_time_to_iso8601(unix_time);
            assert_eq!(expected, got);
        }
    }

    #[test]
    fn test_dates_against_chrono() {
        for year in 1970..=9999 {
            let dt = Utc.with_ymd_and_hms(year, 1, 1, 0, 0, 0).unwrap();
            let expected = dt.to_rfc3339_opts(chrono::SecondsFormat::Secs, true);
            let unix_time = dt.timestamp();
            let got = ExcelDateTime::unix_time_to_iso8601(unix_time as u64);
            assert_eq!(expected, got);

            let dt = Utc.with_ymd_and_hms(year, 1, 2, 0, 0, 0).unwrap();
            let expected = dt.to_rfc3339_opts(chrono::SecondsFormat::Secs, true);
            let unix_time = dt.timestamp();
            let got = ExcelDateTime::unix_time_to_iso8601(unix_time as u64);
            assert_eq!(expected, got);

            let dt = Utc.with_ymd_and_hms(year, 2, 28, 0, 0, 0).unwrap();
            let expected = dt.to_rfc3339_opts(chrono::SecondsFormat::Secs, true);
            let unix_time = dt.timestamp();
            let got = ExcelDateTime::unix_time_to_iso8601(unix_time as u64);
            assert_eq!(expected, got);

            let dt = Utc.with_ymd_and_hms(year, 3, 1, 0, 0, 0).unwrap();
            let expected = dt.to_rfc3339_opts(chrono::SecondsFormat::Secs, true);
            let unix_time = dt.timestamp();
            let got = ExcelDateTime::unix_time_to_iso8601(unix_time as u64);
            assert_eq!(expected, got);

            let dt = Utc.with_ymd_and_hms(year, 12, 31, 0, 0, 0).unwrap();
            let expected = dt.to_rfc3339_opts(chrono::SecondsFormat::Secs, true);
            let unix_time = dt.timestamp();
            let got = ExcelDateTime::unix_time_to_iso8601(unix_time as u64);
            assert_eq!(expected, got);
        }
    }

    #[test]
    fn test_times_against_chrono() {
        for timestamp in 0..=86_4000 {
            let dt = Utc.timestamp_opt(timestamp, 0).unwrap();
            let expected = dt.to_rfc3339_opts(chrono::SecondsFormat::Secs, true);
            let unix_time = dt.timestamp();
            let got = ExcelDateTime::unix_time_to_iso8601(unix_time as u64);
            assert_eq!(expected, got);
        }
    }

    #[test]
    fn test_days_against_chrono() {
        for days in 0..=366 * 4 {
            let timestamp = days * 86_400;
            let dt = Utc.timestamp_opt(timestamp, 0).unwrap();
            let expected = dt.to_rfc3339_opts(chrono::SecondsFormat::Secs, true);
            let unix_time = dt.timestamp();
            let got = ExcelDateTime::unix_time_to_iso8601(unix_time as u64);
            assert_eq!(expected, got);
        }
    }
}
