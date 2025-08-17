use crate::deserialize::Order;
use anyhow::Result;
use chrono::{Datelike, Duration, NaiveDate, NaiveTime, Weekday};
use rust_xlsxwriter::{Color, Format, workbook::Workbook, worksheet::Worksheet};

fn write_date_row(worksheet: &mut Worksheet, index: u32, current_date: f64) -> Result<()> {
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

fn write_order_time(
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

fn write_header_row(worksheet: &mut Worksheet) -> Result<()> {
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

pub fn write_new_xlsx(orders: Vec<Order>) -> Result<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.set_column_width(0, 12)?;
    worksheet.set_column_width(1, 54)?;
    worksheet.set_column_width(2, 45)?;
    worksheet.set_column_width(3, 36)?;
    // count is ok
    worksheet.set_column_width(5, 12)?;
    worksheet.set_column_width(6, 12)?;
    worksheet.set_column_width(7, 12)?;
    worksheet.set_column_width(8, 24)?;

    let mut current_date = orders[0].date;
    let mut index = 0;

    let right_align = Format::new().set_align(rust_xlsxwriter::FormatAlign::Right);

    for order in orders.iter() {
        match index {
            0 => {
                write_header_row(worksheet)?;
                index += 1;

                write_date_row(worksheet, index, current_date)?;
                index += 1;
            }
            _ => {
                if current_date != order.date {
                    current_date = order.date;
                    write_date_row(worksheet, index, order.date)?;
                    index += 1;
                }

                worksheet.write_string(index, 0, order.origin.to_string())?;
                worksheet.write_string(index, 1, order.employee.to_string())?;
                worksheet.write_string(index, 2, order.client.to_string())?;
                worksheet.write_string(index, 3, order.description.to_string())?;
                worksheet.write_string_with_format(
                    index,
                    4,
                    order.count.to_string(),
                    &right_align,
                )?;
                write_order_time(worksheet, index, 5, order.ready, &right_align)?;
                write_order_time(worksheet, index, 6, order.leave, &right_align)?;
                write_order_time(worksheet, index, 7, order.start, &right_align)?;
                worksheet.write_string(index, 8, order.vehicle.to_string())?;

                index += 1;
            }
        }
    }

    workbook.save("./tests/saved.xlsx")?;
    Ok(())
}
