# Convention

The name of a file contains

- T if there is a "Title" column and it is populated
- Tx (for x a number) if there is a title column that uses a different spelling and is populated,
- R if there is a "Researcher" column and it is populated,
- Rx (for x a number) if there is a "Researcher" column that uses a different spelling and is populated,
- A if there is an "Abstract" column and it is populated,
- a if there is an "Abstract" column and it is not populated,
- D if there is a "Doi" column and it is populated,
- d if there is a "Doi" column and it is not populated.
- m if it is some other miscellaneous test.

# Generating the spreadsheet

The spreadsheet is then obtained using

    ssconvert --merge-to=test_input.xlsx *.csv

Installing gnumeric is needed to access this command-line tool.

# Notes

In archives/ are the "problematic" examples:

- m_empty.csv is problematic because of https://github.com/popbr/abstract_doi_finder/issues/13
