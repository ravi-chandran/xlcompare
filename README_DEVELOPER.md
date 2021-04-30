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
```

## Development Iterations
- [Work in development mode](https://packaging.python.org/guides/distributing-packages-using-setuptools/#working-in-development-mode):
- Perform two basic tests.

```bat
cd xlcompare
activate.bat
python -m pip install --editable .
xlcompare -o examples\diffxls.xlsx examples\old.xls examples\new.xls
xlcompare -o examples\diffxlsx.xlsx examples\old.xlsx examples\new.xlsx
```

## Upload To TestPyPI
- Using steps from this [reference](https://packaging.python.org/tutorials/packaging-projects/):

- Create/edit file `.pypirc` in `%HOME%` (Windows) or `$HOME` (Linux):
```
[testpypi]
  username = __token__
  password = enter_the_token_created_on_test.pypi.org
```

```bat
cd xlcompare
activate.bat
python -m build
python -m twine upload --repository testpypi dist/*
```

## Upload To PyPI

## OLD PyPI Notes
- Generate `tar.gz` that will be uploaded with `twine`:
```bat
python setup.py sdist
```



- Push to TestPyPI
```bat
python -m twine upload --repository testpypi dist/*
```

- Push to PyPI
```bat
python -m twine upload dist/*
```

## `pytest` Notes
Tests are written to support both Windows and Linux, although this utility is not really needed in Linux.

- Install `pytest`:
```bat
python -m pip install --upgrade pytest
```

- Local testing:
```bat
pytest -v
```
