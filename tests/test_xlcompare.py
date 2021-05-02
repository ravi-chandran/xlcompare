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

OLD_IDTEST = os.path.join(TESTDIR, 'inputs', 'old_idname.xlsx')
NEW_IDTEST = os.path.join(TESTDIR, 'inputs', 'new_idname.xlsx')


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
    result = verify_common(cmd)
    assert 'Deleted rows: 2' in result.stdout
    assert 'Modified rows: 3' in result.stdout

    # save for visual check, no easy way to automate
    rmfile(SAVEDIFF1)
    os.rename(OUTDIFF, SAVEDIFF1)

    rmfile(OUTDIFF)


# Test comparison of .xlsx vs .xlsx
def test_xlsx_vs_xlsx():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_XLSX, NEW_XLSX, '-o', OUTDIFF]
    result = verify_common(cmd)
    assert 'Deleted rows: 2' in result.stdout
    assert 'Modified rows: 3' in result.stdout

    # save for visual check, no easy way to automate
    rmfile(SAVEDIFF2)
    os.rename(OUTDIFF, SAVEDIFF2)

    rmfile(OUTDIFF)


# Test comparison of .xls vs .xlsx
def test_xls_vs_xlsx():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_XLS, NEW_XLSX, '-o', OUTDIFF]
    result = verify_common(cmd)
    assert 'Deleted rows: 2' in result.stdout
    assert 'Modified rows: 3' in result.stdout
    rmfile(OUTDIFF)


# Test comparison of .xlsx vs .xls
def test_xlsx_vs_xls():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_XLSX, NEW_XLS, '-o', OUTDIFF]
    result = verify_common(cmd)
    assert 'Deleted rows: 2' in result.stdout
    assert 'Modified rows: 3' in result.stdout
    rmfile(OUTDIFF)


# Test comparison of column insertion/deletion
def test_col_change():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_COLS_CHG, NEW_COLS_CHG, '-o', OUTDIFF]
    result = verify_common(cmd)
    assert 'No differences in common columns found' in result.stdout

    # Verify deleted columns - column output not ordered
    assert "Columns in old but not new:" in result.stdout
    assert "'Delete1'" in result.stdout
    assert "'Delete2'" in result.stdout
    assert "'Delete3'" in result.stdout

    # Verify inserted column
    assert "Columns in new but not old: {'Insert 1'}" in result.stdout

    rmfile(OUTDIFF)


# Tests with bad IDs
def test_bad_id_xls():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_XLS, NEW_XLS, '-o', OUTDIFF, '--id', 'BAD_ID']
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert 'ERROR: Column BAD_ID not found in' in result.stdout
    assert result.returncode == 1
    assert not os.path.isfile(OUTDIFF)


def test_bad_id_xlsx():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_XLSX, NEW_XLSX, '-o', OUTDIFF, '--id', 'BAD_ID']
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert 'ERROR: Column BAD_ID not found in' in result.stdout
    assert result.returncode == 1
    assert not os.path.isfile(OUTDIFF)


def test_non_existent_id():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_IDTEST, NEW_IDTEST, '-o', OUTDIFF]
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert 'ERROR: Column ID not found in' in result.stdout
    assert result.returncode == 1
    assert not os.path.isfile(OUTDIFF)


# Test with non-standard good ID
def test_non_standard_good_id():
    rmfile(OUTDIFF)
    cmd = ['xlcompare', OLD_IDTEST, NEW_IDTEST, '-o', OUTDIFF, '--id', 'REQID']
    result = verify_common(cmd)
    assert 'Deleted rows: 2' in result.stdout
    assert 'Modified rows: 3' in result.stdout
    rmfile(OUTDIFF)
