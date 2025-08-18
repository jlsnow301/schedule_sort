use std::env;

use crate::check::test_order_input;
use crate::deserialize::{deserialize_excel, Order};
use crate::errors::{exit_gracefully, throw_windows_err};
use crate::path::get_file_path;
use crate::write::write_new_xlsx;
use anyhow::anyhow;

mod check;
mod deserialize;
mod errors;
mod path;
mod write;

fn main() {
    if env::args().len() != 2 {
        throw_windows_err(anyhow!(
            "This program requires an excel sheet to be dragged onto it."
        ));
        return;
    }

    let args: Vec<String> = env::args().collect();

    let dragged_file = &args[1];

    let file_path = match get_file_path(dragged_file) {
        Err(e) => {
            throw_windows_err(e);
            return;
        }
        Ok(result) => result,
    };

    let mut orders: Vec<Order> = match deserialize_excel(file_path.as_str()) {
        Err(e) => {
            throw_windows_err(e);
            return;
        }
        Ok(result) => result,
    };

    let total = orders.len();

    match test_order_input(&orders) {
        Err(e) => throw_windows_err(e),
        Ok(result) => result,
    };

    orders.sort_by(|a, b| {
        a.date
            .partial_cmp(&b.date)
            .unwrap()
            .then_with(|| a.employee.cmp(&b.employee))
    });

    match write_new_xlsx(orders) {
        Err(e) => throw_windows_err(e),
        Ok(_) => exit_gracefully(format!("Wrote schedule with {} orders", total)),
    }
}
