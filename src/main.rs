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
    }

    let args: Vec<String> = env::args().collect();

    let dragged_file = &args[1];

    let file_path = unwrap_or_throw!(get_file_path(dragged_file));

    let mut orders: Vec<Order> = unwrap_or_throw!(deserialize_excel(file_path.as_str()));

    let total = orders.len();

    unwrap_or_throw!(test_order_input(&orders));

    orders.sort_by(|a, b| {
        a.date
            .partial_cmp(&b.date)
            .unwrap()
            .then_with(|| a.employee.cmp(&b.employee))
            .then_with(|| a.ready.partial_cmp(&b.ready).unwrap())
    });

    match write_new_xlsx(orders) {
        Err(e) => throw_windows_err(e),
        Ok(_) => exit_gracefully(format!("Wrote schedule with {} orders", total)),
    }
}
