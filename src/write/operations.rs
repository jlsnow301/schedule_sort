use crate::{
    deserialize::Order,
    write::util::{
        write_daily_count_sum, write_date_row, write_final_row, write_header_row, write_order_time,
    },
};
use anyhow::Result;
use rust_xlsxwriter::{Color, Format, workbook::Workbook};

pub fn write_new_xlsx(orders: Vec<Order>) -> Result<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

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
    let mut index = 0;
    let mut daily_orders = 0;
    let mut sum_rows: Vec<u32> = Vec::new();

    let bold = Format::new().set_bold();
    let eastlake = Format::new().set_font_color(Color::RGB(0xBE5014));
    let empty = Format::new();
    let fremont = Format::new().set_font_color(Color::Green);
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
                    write_daily_count_sum(
                        worksheet,
                        index,
                        &mut daily_orders,
                        &mut sum_rows,
                        &bold,
                    )?;
                    write_date_row(worksheet, index + 2, order.date)?;
                    index += 3;
                }

                let to_use = match order.origin.as_str() {
                    "Fremont" => &fremont,
                    "Eastlake" => &eastlake,
                    _ => &empty,
                };

                worksheet.write_string_with_format(index, 0, order.origin.to_string(), to_use)?;
                worksheet.write_string(index, 1, order.employee.to_string())?;
                worksheet.write_string(index, 2, order.client.to_string())?;
                worksheet.write_string(index, 3, order.description.to_string())?;
                worksheet.write_number_with_format(index, 4, order.count as f64, &right_align)?;
                write_order_time(worksheet, index, 5, order.ready, &right_align)?;
                write_order_time(worksheet, index, 6, order.leave, &right_align)?;
                write_order_time(worksheet, index, 7, order.start, &right_align)?;
                worksheet.write_string(index, 8, order.vehicle.to_string())?;

                index += 1;
                daily_orders += 1;
            }
        }
    }

    write_daily_count_sum(worksheet, index, &mut daily_orders, &mut sum_rows, &bold)?;
    write_final_row(worksheet, index, &sum_rows, orders.len() as u32)?;

    workbook.save("./formatted.xlsx")?;
    Ok(())
}
