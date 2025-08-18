use anyhow::{Context, Result};
use calamine::{open_workbook, Data, DataType, ExcelDateTime, Reader, Xlsx};

#[derive(Debug)]
pub struct Order {
    pub date: f64,
    pub origin: String,
    pub employee: String,
    pub client: String,
    pub description: String,
    pub count: i64,
    pub ready: f64,
    pub leave: f64,
    pub start: f64,
    pub vehicle: String,
}

fn deserialize_string_cell(cell: Option<&Data>, default: &str) -> String {
    cell.and_then(|cell| cell.get_string())
        .unwrap_or(default)
        .to_string()
}

fn deserialize_date_cell(cell: Option<&Data>, default: f64) -> f64 {
    let excel_dt = cell
        .and_then(|cell| {
            if let Data::DateTime(excel_datetime) = cell {
                Some(*excel_datetime)
            } else {
                None
            }
        })
        .unwrap_or(ExcelDateTime::new(
            default,
            calamine::ExcelDateTimeType::DateTime,
            false,
        ));

    excel_dt.as_f64()
}

pub fn deserialize_excel(file_path: &str) -> Result<Vec<Order>> {
    let mut workbook: Xlsx<_> = open_workbook(file_path)
        .with_context(|| format!("Failed to open Excel file: {}", file_path))?;

    let worksheet = workbook
        .worksheet_range_at(0)
        .context("Cannot find worksheet at index 0")?
        .context("Error reading worksheet data")?;

    let mut orders = Vec::new();

    let length = worksheet.height();
    for (index, row) in worksheet.rows().enumerate() {
        match index {
            0 => continue,
            i if i >= length - 2 => break,
            _ => {
                let order = Order {
                    date: deserialize_date_cell(row.first(), 45658.0),
                    origin: deserialize_string_cell(row.get(1), ""),
                    employee: deserialize_string_cell(row.get(2), ""),
                    client: deserialize_string_cell(row.get(3), ""),
                    description: deserialize_string_cell(row.get(4), ""),
                    count: row
                        .get(5)
                        .map(|x| match x {
                            Data::String(x) => x.parse::<i64>().unwrap_or(0),
                            Data::Int(x) => *x,
                            _ => 0,
                        })
                        .unwrap(),
                    ready: deserialize_date_cell(row.get(6), 0.0),
                    leave: deserialize_date_cell(row.get(7), 0.0),
                    start: deserialize_date_cell(row.get(8), 0.0),
                    vehicle: deserialize_string_cell(row.get(9), ""),
                };

                orders.push(order);
            }
        }
    }
    Ok(orders)
}
