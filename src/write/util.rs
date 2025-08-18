use anyhow::Result;
use chrono::{Datelike, Duration, NaiveDate, NaiveTime, Weekday};
use rust_xlsxwriter::{Color, Format, worksheet::Worksheet};

pub fn write_date_row(worksheet: &mut Worksheet, index: u32, current_date: f64) -> Result<()> {
    let start = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
    let days_to_add = current_date as i64;
    let date = start
        .checked_add_signed(Duration::days(days_to_add))
        .unwrap();

    let day_of_week = date.format("%A").to_string();
    let formatted_date = date.format("%m/%d/%Y").to_string();
    let date_string = format!("{}, {}", day_of_week, formatted_date);

    let bg_color = match date.weekday() {
        Weekday::Mon => Color::RGB(0xFFB3BA),
        Weekday::Tue => Color::RGB(0xFFDFBA),
        Weekday::Wed => Color::RGB(0xFFFFBA),
        Weekday::Thu => Color::RGB(0xBAFFC9),
        Weekday::Fri => Color::RGB(0xBAE1FF),
        Weekday::Sat => Color::RGB(0xC9BAFF),
        Weekday::Sun => Color::RGB(0xFFBAF3),
    };

    let format = rust_xlsxwriter::Format::new().set_background_color(bg_color);

    worksheet.merge_range(index, 0, index, 8, date_string.as_str(), &format)?;

    Ok(())
}

pub fn write_order_time(
    worksheet: &mut Worksheet,
    index: u32,
    col: u16,
    excel_dt: f64,
    format: &Format,
) -> Result<()> {
    if excel_dt == 0.0 {
        return Ok(());
    }

    let time_fraction = excel_dt - excel_dt.floor();
    let total_seconds = (time_fraction * 86400.0).round() as u32;

    let hours = total_seconds / 3600;
    let minutes = (total_seconds % 3600) / 60;
    let seconds = total_seconds % 60;

    let date = NaiveTime::from_hms_opt(hours, minutes, seconds)
        .unwrap()
        .format("%I:%M %p")
        .to_string();

    worksheet.write_string_with_format(index, col, date, format)?;

    Ok(())
}

pub fn write_header_row(worksheet: &mut Worksheet) -> Result<()> {
    let headers = Format::new().set_bold();

    worksheet.write_row_with_format(
        0,
        0,
        vec![
            "Origin",
            "Employee",
            "Client",
            "Description",
            "Count",
            "Ready",
            "Leave",
            "Start",
            "Vehicle",
        ],
        &headers,
    )?;

    Ok(())
}

pub fn write_daily_count_sum(
    worksheet: &mut Worksheet,
    index: u32,
    daily_orders: &mut u32,
    sum_rows: &mut Vec<u32>,
    format: &Format,
) -> Result<()> {
    // Starts at 1 :)
    let excel_index = index + 1;
    let first_order = excel_index - *daily_orders;

    worksheet.write_string_with_format(index, 1, format!("{} orders", *daily_orders), format)?;
    worksheet.write_formula_with_format(
        index,
        4,
        format!("=SUM(E{}:E{})", first_order, index).as_str(),
        format,
    )?;

    sum_rows.push(excel_index);
    *daily_orders = 0;

    Ok(())
}

pub fn write_final_row(
    worksheet: &mut Worksheet,
    index: u32,
    sum_rows: &[u32],
    total_orders: u32,
) -> Result<()> {
    let sums = sum_rows
        .iter()
        .map(|n| format!("E{}", n))
        .collect::<Vec<String>>()
        .join(",");

    let last = index + 2;
    let mut vec_string = vec![String::new(); 9];
    vec_string[0] = "Totals".to_string();
    vec_string[1] = format!("{} orders", total_orders);

    let last_sum = Format::new()
        .set_bold()
        .set_background_color(Color::RGB(0xF5F5F5));

    worksheet.write_row_with_format(last, 0, vec_string, &last_sum)?;
    worksheet.write_formula_with_format(last, 4, format!("=SUM({})", sums).as_str(), &last_sum)?;

    Ok(())
}
