#!/usr/bin/env python3
import filecmp
import os
import pytest
import subprocess


TESTDIR = os.path.dirname(os.path.realpath(__file__))

OLD_XLS = os.path.join(TESTDIR, 'inputs', 'old.xls')
NEW_XLS = os.path.join(TESTDIR, 'inputs', 'new.xls')

OLD_XLSX = os.path.join(TESTDIR, 'inputs', 'old.xlsx')
NEW_XLSX = os.path.join(TESTDIR, 'inputs', 'new.xlsx')

EXPECTED = os.path.join(TESTDIR, 'expected', 'diffxls.xlsx')

OUTDIFF = os.path.join(TESTDIR, 'diff.xlsx')
SAVEDIFF1 = os.path.join(TESTDIR, 'save1.xlsx')
SAVEDIFF2 = os.path.join(TESTDIR, 'save2.xlsx')


def rmfile(filepath):
    """Remove file if it exists."""
    if os.path.isfile(filepath):
        os.remove(filepath)


def compare_files(file1, file2):
    """Compare two binary files."""
    with open(file1, 'rb') as f:
        data1 = f.read()
    with open(file2, 'rb') as f:
        data2 = f.read()
    return data1 == data2


def compare(cmd):
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert result.stderr == ''
    assert 'Generated' in result.stdout
    assert 'Done.' in result.stdout
    assert os.path.isfile(OUTDIFF)
    # assert compare_files(OUTDIFF, EXPECTED)
    # filecmp.clear_cache()
    # assert filecmp.cmp(OUTDIFF, EXPECTED)


# Test basic entry point to xlcompare
def test_entrypoint():
    exit_status = os.system('xlcompare --help')
    assert exit_status == 0


# Test comparison of .xls vs .xls
def test_xls_vs_xls():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_XLS, NEW_XLS, '-o', OUTDIFF]
    compare(cmd)

    # save for visual check, no easy way to automate
    rmfile(SAVEDIFF1)
    os.rename(OUTDIFF, SAVEDIFF1)

    rmfile(OUTDIFF)


# Test comparison of .xlsx vs .xlsx
def test_xlsx_vs_xlsx():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_XLSX, NEW_XLSX, '-o', OUTDIFF]
    compare(cmd)

    # save for visual check, no easy way to automate
    rmfile(SAVEDIFF2)
    os.rename(OUTDIFF, SAVEDIFF2)

    rmfile(OUTDIFF)


# Test comparison of .xls vs .xlsx
def test_xls_vs_xlsx():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_XLS, NEW_XLSX, '-o', OUTDIFF]
    compare(cmd)
    rmfile(OUTDIFF)


# Test comparison of .xlsx vs .xls
def test_xlsx_vs_xls():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_XLSX, NEW_XLS, '-o', OUTDIFF]
    compare(cmd)
    rmfile(OUTDIFF)

# TODO:
# Test for --id
# Test for inserted column
# Test for deleted column
# Feature addition: report statistics
