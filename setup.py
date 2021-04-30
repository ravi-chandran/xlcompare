#!/usr/bin/env python3
import setuptools

with open("README.md", "r") as f:
    long_description = f.read()

setuptools.setup(
    name="xlcompare",
    version="0.1.0",
    packages=setuptools.find_packages(),

    entry_points = {
        'console_scripts': [
            'xlcompare=xlcompare.xlcompare:main'
        ],
    },

    python_requires=">=3.6",
    install_requires=[
        "xlrd>=2.0.1",
        "pylightxl>=1.54",
        "XlsxWriter>=1.3.9"
    ],

    author="Ravi Chandran",
    description="Compare two Excel files where rows have unique identifiers.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/ravi-chandran/xlcompare",

    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)