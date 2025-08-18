use crate::deserialize::Order;
use anyhow::{anyhow, Result};

pub fn test_order_input(orders: &[Order]) -> Result<()> {
    // No orders
    if orders.is_empty() {
        return Err(anyhow!("No orders found in Excel file"));
    }

    // Missing first date
    let mut current_date = orders[0].date;
    if current_date == 45658.0 {
        return Err(anyhow!("First order date is invalid"));
    }

    // Missing valid dates
    let mut days_total = 0;
    for order in orders.iter() {
        if order.date == 45658.0 {
            return Err(anyhow!("Invalid order date found"));
        }
        if order.date != current_date {
            days_total += 1;
            current_date = order.date;
        }
    }
    if days_total == 0 {
        return Err(anyhow!("No valid dates found"));
    }

    Ok(())
}
