use crate::{
    deserialize::Order,
    write::util::{
        write_daily_count_sum, write_date_row, write_final_row, write_header_row, write_order_time,
    },
};
use anyhow::Result;
use rust_xlsxwriter::{workbook::Workbook, Color, Format, FormatAlign, FormatBorder};

const LT_GRAY: u32 = 0xE5E7EB;
const RUST: u32 = 0xBE5014; // crab

pub fn write_new_xlsx(orders: Vec<Order>) -> Result<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Default sizing
    worksheet.set_column_width(0, 12)?;
    worksheet.set_column_width(1, 45)?;
    worksheet.set_column_width(2, 45)?;
    worksheet.set_column_width(3, 36)?;
    // count is ok
    worksheet.set_column_width(5, 12)?;
    worksheet.set_column_width(6, 12)?;
    worksheet.set_column_width(7, 12)?;
    worksheet.set_column_width(8, 24)?;

    let mut current_date = orders[0].date;
    let mut row = 0;
    let mut daily_orders = 0;
    let mut sum_rows: Vec<u32> = Vec::new();

    // Create themes
    let standard = Format::new()
        .set_bold()
        .set_border(FormatBorder::Thin)
        .set_border_color(Color::Gray);

    let eastlake = standard.clone().set_font_color(Color::RGB(RUST));
    let last_sum = Format::new()
        .set_bold()
        .set_background_color(Color::RGB(LT_GRAY));
    let fremont = standard.clone().set_font_color(Color::Green);
    let header = standard.clone().set_background_color(Color::RGB(LT_GRAY));
    let right_align = standard.clone().set_align(FormatAlign::Right);

    write_date_row(worksheet, row, current_date)?;
    row += 1;
    write_header_row(worksheet, row, &header)?;
    row += 1;

    for order in orders.iter() {
        // Writes daily counts followed by a new date
        if current_date != order.date {
            current_date = order.date;
            write_daily_count_sum(worksheet, row, &mut daily_orders, &mut sum_rows, &standard)?;
            // We want a blank row after the counts
            write_date_row(worksheet, row + 2, order.date)?;
            write_header_row(worksheet, row + 3, &header)?;
            row += 4;
        };

        let to_use = match order.origin.as_str() {
            "Fremont" => &fremont,
            "Eastlake" => &eastlake,
            _ => &standard,
        };

        worksheet.set_row_format(row, &standard)?;
        worksheet.write_string_with_format(row, 0, order.origin.to_string(), to_use)?;
        worksheet.write_string(row, 1, order.employee.to_string())?;
        worksheet.write_string(row, 2, order.client.to_string())?;
        worksheet.write_string(row, 3, order.description.to_string())?;
        worksheet.write_number_with_format(row, 4, order.count as f64, &right_align)?;
        write_order_time(worksheet, row, 5, order.ready, &right_align)?;
        write_order_time(worksheet, row, 6, order.leave, &right_align)?;
        write_order_time(worksheet, row, 7, order.start, &right_align)?;
        worksheet.write_string(row, 8, order.vehicle.to_string())?;

        row += 1;
        daily_orders += 1;
    }

    write_daily_count_sum(worksheet, row, &mut daily_orders, &mut sum_rows, &standard)?;
    write_final_row(worksheet, row, &sum_rows, orders.len() as u32, &last_sum)?;

    workbook.save("./formatted_schedule.xlsx")?;
    Ok(())
}
