#![feature(specialization, const_fn)]
extern crate pyo3;

use pyo3::prelude::*;
use pyo3::exceptions;

extern crate calamine;

use calamine::{Sheets, Range, DataType, Reader};

fn to_py_err(err: calamine::Error) -> pyo3::PyErr
{
    PyErr::new::<exceptions::ValueError, _>(format!("{}", err).to_string())
}

#[pyclass]
struct Workbook {
    sheets: Sheets,
}

#[pyclass]
struct Worksheet {
    range: Range<DataType>,
}

#[pymethods]
impl Workbook {
    #[new]
    fn __new__(obj: &PyRawObject, path: String) {
        obj.init({
            Workbook {
                sheets: calamine::open_workbook_auto(path).expect("Cannot open file"),
            }
        });
    }

    fn sheet_names(&mut self) -> PyResult<Vec<String>> {
        Ok(self.sheets.sheet_names().to_vec())
    }

    fn get_sheet(&mut self, name: String, py: Python) -> PyResult<Py<Worksheet>> {
        let range = self.sheets.worksheet_range(name.as_str()).unwrap_or_else(||
            Err(calamine::Error::Msg("sheet not found"))).map_err(to_py_err)?;
        Py::new(py, Worksheet { range, })
    }
}

impl Worksheet {
    fn _get_size(&self) -> (u32, u32) {
        self.range.end().map_or_else(
            || (0, 0),
            |v| (v.0 + 1, v.1 + 1)
        )
    }
}

#[pymethods]
impl Worksheet {
    fn get_size(&self) -> PyResult<(u32, u32)> {
        Ok(self._get_size())
    }

    fn width(&self) -> PyResult<u32> {
        Ok(self._get_size().1)
    }

    fn height(&self) -> PyResult<u32> {
        Ok(self._get_size().0)
    }

    fn get_value(&self, row: u32, col: u32, py: Python) -> PyResult<pyo3::PyObject> {
        match self.range.get_value((row, col)) {
            None => { Ok(().to_object(py)) }
            Some(calamine::DataType::Int(i)) => { Ok(i.to_object(py)) }
            Some(calamine::DataType::Float(i)) => { Ok(i.to_object(py)) }
            Some(calamine::DataType::String(ref i)) => { Ok(i.clone().to_object(py)) }
            Some(calamine::DataType::Bool(i)) => { Ok(i.to_object(py)) }
            Some(calamine::DataType::Empty) => { Ok(().to_object(py)) }
            Some(calamine::DataType::Error(ref e)) => {
                match e {
                    &calamine::CellErrorType::Div0 => {
                        Err(PyErr::new::<exceptions::ValueError, _>("Division by 0 error"))
                    }
                    &calamine::CellErrorType::NA => {
                        Err(PyErr::new::<exceptions::ValueError, _>("Unavailable value error"))
                    }
                    &calamine::CellErrorType::Name => {
                        Err(PyErr::new::<exceptions::ValueError, _>("Invalid name error"))
                    }
                    &calamine::CellErrorType::Null => {
                        Err(PyErr::new::<exceptions::ValueError, _>("Null value error"))
                    }
                    &calamine::CellErrorType::Num => {
                        Err(PyErr::new::<exceptions::ValueError, _>("Number error"))
                    }
                    &calamine::CellErrorType::Ref => {
                        Err(PyErr::new::<exceptions::ValueError, _>("Invalid cell reference error"))
                    }
                    &calamine::CellErrorType::Value => {
                        Err(PyErr::new::<exceptions::ValueError, _>("Value error"))
                    }
                    &calamine::CellErrorType::GettingData => {
                        Err(PyErr::new::<exceptions::ValueError, _>("Getting data"))
                    }
                }
            }
        }
    }
}

// add bindings to the generated python module
// N.B: names: "librust2py" must be the name of the `.so` or `.pyd` file
/// This module is implemented in Rust.
#[pymodule]
fn pyxlsx_rs(_py: Python, m: &PyModule) -> PyResult<()> {
  m.add_class::<Workbook>()?;
  m.add_class::<Worksheet>()?;

  Ok(())
}

