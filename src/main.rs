use std::env;

use crate::check::test_order_input;
use crate::deserialize::{deserialize_excel, Order};
use crate::errors::{exit_gracefully, throw_windows_err};
use crate::path::get_file_path;
use crate::write::write_new_xlsx;
use anyhow::{anyhow, Result};

mod check;
mod deserialize;
mod errors;
mod path;
mod write;

fn sort_orders(orders: &mut [Order]) -> Result<()> {
    orders.sort_by(|a, b| {
        a.date
            .partial_cmp(&b.date)
            .unwrap()
            .then_with(|| a.employee.cmp(&b.employee))
            .then_with(|| a.ready.partial_cmp(&b.ready).unwrap())
    });
    Ok(())
}

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

    unwrap_or_throw!(sort_orders(&mut orders));

    match write_new_xlsx(orders, "./formatted_schedule.xlsx") {
        Err(e) => throw_windows_err(e),
        Ok(_) => exit_gracefully(format!("Wrote schedule with {} orders", total)),
    }
}

#[cfg(test)]
mod tests {
    use calamine::{open_workbook, Reader, Xlsx};

    use super::*;

    #[test]
    fn it_works() {
        let mut orders =
            deserialize_excel("./tests/mock_data.xlsx").expect("Couldn't deserialize file");

        sort_orders(&mut orders).expect("Couldn't sort orders");

        write_new_xlsx(orders, "./tests/formatted_test.xlsx")
            .expect("Couldn't write the new sheet");

        let mut workbook: Xlsx<_> = open_workbook("./tests/formatted_test.xlsx")
            .expect("Couldn't open formatted excel sheet");

        let worksheet = workbook
            .worksheet_range_at(0)
            .expect("Cannot find worksheet at index 0")
            .expect("Error reading worksheet data");

        let height = worksheet.rows().len();
        assert_eq!(height, 79);

        let first_driver = worksheet.get((2, 1)).expect("Couldn't get first driver");
        assert_eq!(first_driver, "Alfonse Laven");
    }
}
