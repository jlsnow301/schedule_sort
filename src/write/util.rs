use anyhow::Result;
use chrono::{Datelike, Duration, NaiveDate, NaiveTime, Weekday};
use rust_xlsxwriter::{worksheet::Worksheet, Color, Format, FormatAlign};

const PASTEL_RED: u32 = 0xFFB3BA;
const PASTEL_ORANGE: u32 = 0xFFDFBA;
const PASTEL_YELLOW: u32 = 0xFFFFBA;
const PASTEL_GREEN: u32 = 0xBAFFC9;
const PASTEL_BLUE: u32 = 0xBAE1FF;
const PASTEL_PURPLE: u32 = 0xC9BAFF;
const PASTEL_PINK: u32 = 0xFFBAF3;

pub fn write_date_row(worksheet: &mut Worksheet, row: u32, current_date: f64) -> Result<()> {
    let start = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
    let days_to_add = current_date as i64;
    let date = start
        .checked_add_signed(Duration::days(days_to_add))
        .unwrap();

    let day_of_week = date.format("%A").to_string();
    let formatted_date = date.format("%m/%d/%Y").to_string();
    let date_string = format!("{}, {}", day_of_week, formatted_date);

    let bg_color = match date.weekday() {
        Weekday::Mon => Color::RGB(PASTEL_RED),
        Weekday::Tue => Color::RGB(PASTEL_ORANGE),
        Weekday::Wed => Color::RGB(PASTEL_YELLOW),
        Weekday::Thu => Color::RGB(PASTEL_GREEN),
        Weekday::Fri => Color::RGB(PASTEL_BLUE),
        Weekday::Sat => Color::RGB(PASTEL_PURPLE),
        Weekday::Sun => Color::RGB(PASTEL_PINK),
    };

    let format = Format::new().set_background_color(bg_color);

    worksheet.merge_range(row, 0, row, 8, date_string.as_str(), &format)?;

    Ok(())
}

pub fn write_order_time(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    excel_dt: f64,
    format: &Format,
) -> Result<()> {
    if excel_dt == 0.0 {
        worksheet.write_string_with_format(row, col, "", format)?;
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

    worksheet.write_string_with_format(row, col, date, format)?;

    Ok(())
}

pub fn write_header_row(worksheet: &mut Worksheet, row: u32, format: &Format) -> Result<()> {
    worksheet.set_row_format(row, format)?;
    worksheet.write_row(row, 0, vec!["Origin", "Employee", "Client", "Description"])?;
    worksheet.write_row_with_format(
        row,
        4,
        vec!["Count", "Ready", "Leave", "Start"],
        &format.clone().set_align(FormatAlign::Right),
    )?;
    worksheet.write_string(row, 8, "Vehicle")?;

    Ok(())
}

pub fn write_daily_count_sum(
    worksheet: &mut Worksheet,
    row: u32,
    daily_orders: &mut u32,
    sum_rows: &mut Vec<u32>,
    format: &Format,
) -> Result<()> {
    // Starts at 1 :)
    let excel_index = row + 1;
    let first_order = excel_index - *daily_orders;

    worksheet.set_row_format(row, format)?;
    worksheet.write_string(row, 1, format!("{} orders", *daily_orders))?;
    worksheet.write_formula(row, 4, format!("=SUM(E{}:E{})", first_order, row).as_str())?;

    sum_rows.push(excel_index);
    *daily_orders = 0;

    Ok(())
}

pub fn write_final_row(
    worksheet: &mut Worksheet,
    row: u32,
    sum_rows: &[u32],
    total_orders: u32,
    format: &Format,
) -> Result<()> {
    let sums = sum_rows
        .iter()
        .map(|n| format!("E{}", n))
        .collect::<Vec<String>>()
        .join(",");

    let last = row + 2;
    let mut vec_string = vec![String::new(); 9];
    vec_string[0] = "Totals".to_string();
    vec_string[1] = format!("{} orders", total_orders);

    worksheet.set_row_format(last, format)?;
    worksheet.write_row(last, 0, vec_string)?;
    worksheet.write_formula(last, 4, format!("=SUM({})", sums).as_str())?;

    Ok(())
}
