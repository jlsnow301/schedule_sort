use crate::deserialize::{Order, deserialize_excel};
use crate::write::write_new_xlsx;
use anyhow::{Context, Result};

mod deserialize;
mod write;

fn main() -> Result<()> {
    let path = "./tests/current.xlsx";

    let mut orders: Vec<Order> = deserialize_excel(path).context("Could not deserialize excel")?;

    // No orders
    if orders.is_empty() {
        return Err(anyhow::anyhow!("No orders found in the Excel file"));
    }

    // Missing first date
    let mut current_date = orders[0].date;
    if current_date == 45658.0 {
        return Err(anyhow::anyhow!("First order date is not valid"));
    }

    // Missing valid dates
    let mut days_total = 0;
    for order in orders.iter() {
        if order.date != current_date {
            days_total += 1;
            current_date = order.date;
        }
    }
    if days_total == 0 {
        return Err(anyhow::anyhow!("No valid dates found in the orders"));
    }

    let total = orders.len();

    orders.sort_by(|a, b| {
        a.date
            .partial_cmp(&b.date)
            .unwrap()
            .then_with(|| a.employee.cmp(&b.employee))
    });

    write_new_xlsx(orders).context("Could not write new excel sheet")?;

    println!("Excel file written with {} orders", total);
    Ok(())
}
