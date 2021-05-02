#!/usr/bin/env python3
import os
import pytest
import subprocess


TESTDIR = os.path.dirname(os.path.realpath(__file__))

OLD_XLS = os.path.join(TESTDIR, 'inputs', 'old.xls')
NEW_XLS = os.path.join(TESTDIR, 'inputs', 'new.xls')

OLD_XLSX = os.path.join(TESTDIR, 'inputs', 'old.xlsx')
NEW_XLSX = os.path.join(TESTDIR, 'inputs', 'new.xlsx')

OLD_COLS_CHG = os.path.join(TESTDIR, 'inputs', 'old_columns_change.xlsx')
NEW_COLS_CHG = os.path.join(TESTDIR, 'inputs', 'new_columns_change.xlsx')

# EXPECTED = os.path.join(TESTDIR, 'expected', 'diffxls.xlsx')

OUTDIFF = os.path.join(TESTDIR, 'diff.xlsx')
SAVEDIFF1 = os.path.join(TESTDIR, 'save1.xlsx')
SAVEDIFF2 = os.path.join(TESTDIR, 'save2.xlsx')


def rmfile(filepath):
    """Remove file if it exists."""
    if os.path.isfile(filepath):
        os.remove(filepath)


def verify_common(cmd):
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert result.stderr == ''
    assert 'Generated' in result.stdout
    assert 'Done.' in result.stdout
    assert os.path.isfile(OUTDIFF)
    return result


# Test basic entry point to xlcompare
def test_entrypoint():
    exit_status = os.system('xlcompare --help')
    assert exit_status == 0


# Test comparison of .xls vs .xls
def test_xls_vs_xls():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_XLS, NEW_XLS, '-o', OUTDIFF]
    verify_common(cmd)

    # save for visual check, no easy way to automate
    rmfile(SAVEDIFF1)
    os.rename(OUTDIFF, SAVEDIFF1)

    rmfile(OUTDIFF)


# Test comparison of .xlsx vs .xlsx
def test_xlsx_vs_xlsx():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_XLSX, NEW_XLSX, '-o', OUTDIFF]
    verify_common(cmd)

    # save for visual check, no easy way to automate
    rmfile(SAVEDIFF2)
    os.rename(OUTDIFF, SAVEDIFF2)

    rmfile(OUTDIFF)


# Test comparison of .xls vs .xlsx
def test_xls_vs_xlsx():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_XLS, NEW_XLSX, '-o', OUTDIFF]
    verify_common(cmd)
    rmfile(OUTDIFF)


# Test comparison of .xlsx vs .xls
def test_xlsx_vs_xls():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_XLSX, NEW_XLS, '-o', OUTDIFF]
    verify_common(cmd)
    rmfile(OUTDIFF)


# Test comparison of column insertion/deletion
def test_col_change():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_COLS_CHG, NEW_COLS_CHG, '-o', OUTDIFF]
    result = verify_common(cmd)

    # Verify deleted columns - column output not ordered
    assert "Columns in old but not new:" in result.stdout
    assert "'Delete1'" in result.stdout
    assert "'Delete2'" in result.stdout
    assert "'Delete3'" in result.stdout

    # Verify inserted column
    assert "Columns in new but not old: {'Insert 1'}" in result.stdout

    rmfile(OUTDIFF)




# TODO:
# Test for --id
# Feature addition: report statistics
