#![feature(proc_macro, specialization, const_fn)]
extern crate pyo3;
use pyo3::prelude::*;

extern crate calamine;
use calamine::{Sheets, Range, DataType};

fn to_py_err(err: calamine::Error) -> pyo3::PyErr
{
    PyErr::new::<exc::ValueError, _>(err.description().to_string())
}

#[py::class]
struct Workbook {
    sheets: Sheets,
    token: PyToken,
}

#[py::class]
struct Worksheet {
    range: Range<DataType>,
    token: PyToken,
}

#[py::methods]
impl Workbook {
    #[new]
    fn __new__(obj: &PyRawObject, path: String) -> PyResult<()> {
        obj.init(|token| {
            Workbook {
                sheets: Sheets::open(path).expect("Cannot open file"),
                token: token
            }
        })
    }

    fn sheet_names(&mut self) -> PyResult<Vec<String>> {
        self.sheets.sheet_names().map_err(to_py_err)
    }

    fn get_sheet(&mut self, name: String, py: Python) -> PyResult<Py<Worksheet>> {
        let range = self.sheets.worksheet_range(name.as_str()).map_err(to_py_err)?;
        py.init(|token| {
            Worksheet {
                range: range,
                token: token
            }
        })
    }
}

#[py::methods]
impl Worksheet {
    fn get_size(&self) -> PyResult<(usize, usize)> {
        Ok(self.range.get_size())
    }

    fn width(&self) -> PyResult<usize> {
        Ok(self.range.width())
    }

    fn height(&self) -> PyResult<usize> {
        Ok(self.range.height())
    }

    fn get_value(&self, row: usize, col: usize, py: Python) -> PyResult<pyo3::PyObject> {
        if row >= self.range.height() {
            return Err(PyErr::new::<exc::IndexError, _>("width out of bound"))
        }
        if col >= self.range.width() {
            return Err(PyErr::new::<exc::IndexError, _>("height out of bound"))
        }
        match *self.range.get_value((row as u32, col as u32)) {
           calamine::DataType::Int(i) => { Ok(i.into_object(py)) }
           calamine::DataType::Float(i) => { Ok(i.into_object(py)) }
           calamine::DataType::String(ref i) => { Ok(i.clone().into_object(py)) }
           calamine::DataType::Bool(i) => { Ok(i.into_object(py)) }
           calamine::DataType::Empty => { Ok(().into_object(py)) }
           calamine::DataType::Error(ref e) => {
                match e {
                    &calamine::CellErrorType::Div0 => {
                        Err(PyErr::new::<exc::ValueError, _>("Division by 0 error"))
                    }
                    &calamine::CellErrorType::NA => {
                        Err(PyErr::new::<exc::ValueError, _>("Unavailable value error"))
                    }
                    &calamine::CellErrorType::Name => {
                        Err(PyErr::new::<exc::ValueError, _>("Invalid name error"))
                    }
                    &calamine::CellErrorType::Null => {
                        Err(PyErr::new::<exc::ValueError, _>("Null value error"))
                    }
                    &calamine::CellErrorType::Num => {
                        Err(PyErr::new::<exc::ValueError, _>("Number error"))
                    }
                    &calamine::CellErrorType::Ref => {
                        Err(PyErr::new::<exc::ValueError, _>("Invalid cell reference error"))
                    }
                    &calamine::CellErrorType::Value => {
                        Err(PyErr::new::<exc::ValueError, _>("Value error"))
                    }
                    &calamine::CellErrorType::GettingData => {
                        Err(PyErr::new::<exc::ValueError, _>("Getting data"))
                    }
                }
            }
        }
    }
}

// add bindings to the generated python module
// N.B: names: "librust2py" must be the name of the `.so` or `.pyd` file
/// This module is implemented in Rust.
#[py::modinit(pyxlsx_rs)]
fn init_mod(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_class::<Workbook>()?;
    m.add_class::<Worksheet>()?;

    Ok(())
}

