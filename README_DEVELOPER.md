# Developer Notes

## Libraries Used
- I don't like `openpyxl`. It's kinda clunky and slower than `XlsxWriter`. Since `xlcompare` does not need to read and write to the same file, `openpyxl` is unnecessary.
- `xlrd`: I've used this library and it works well for `.xls`. It used to also work for `.xlsx` but newer versions don't support it any longer.
- `pylightxl`: Something new I haven't tried before. Works as a great light weight `.xlsx` file reader.
- `XlsxWriter`: Great for writing `.xlsx` files.


## Virtual Environment Setup for Development
```bat
cd xlcompare
python -m venv venv1
activate.bat
venv1\Scripts\python -m pip install --upgrade pip
venv1\Scripts\python -m pip install --upgrade setuptools
venv1\Scripts\python -m pip install --upgrade build
venv1\Scripts\python -m pip install --upgrade twine
venv1\Scripts\python -m pip install xlrd pylightxl XlsxWriter
venv1\Scripts\python -m pip install --upgrade pycodestyle
venv1\Scripts\python -m pip install --upgrade pytest
```

## Development Iterations
- [Work in development mode](https://packaging.python.org/guides/distributing-packages-using-setuptools/#working-in-development-mode):
- Perform two basic tests.

```bat
cd xlcompare
activate.bat
python -m pip install --editable .
pytest -v
```

## Configure TestPyPI and PyPI Access
- Using steps from this [reference](https://packaging.python.org/tutorials/packaging-projects/):

- Create/edit file `.pypirc` in `%HOME%` (Windows) or `$HOME` (Linux):
```
[testpypi]
  username = __token__
  password = enter the token created on test.pypi.org

[pypi]
  username = __token__
  password = enter the token created on pypi.org
```

## Upload To TestPyPI
- Bump version in `setup.py`
- Build and upload
```bat
cd xlcompare
activate.bat
python -m build
python -m twine upload --repository testpypi dist/*
```

- Install and try it out in a separate `venv`.

- Note that TestPyPI provides the following install. However, the dependencies (`xlrd>=2.0.1` etc) cannot be found at `test.pypi.org` and will need to be installed manually.
```bat
pip install -i https://test.pypi.org/simple/ xlcompare
```

## Upload To PyPI
- Bump version in `setup.py`
- Build and upload
```bat
cd xlcompare
activate.bat
python -m build
python -m twine upload dist/*
```
