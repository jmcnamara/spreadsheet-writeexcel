NAME
    Spreadsheet::WriteExcel - Write text and numbers to minimal
    Excel binary file.

DESCRIPTION
    This module can be used to write numbers and text in the native
    Excel binary file format. This is a minimal implementation of an
    Excel file; no formatting can be applied to cells and only a
    single worksheet can be written to a workbook.

    It is intended to be cross-platform, however, this is not
    guaranteed. See the section on portability in the main
    documentation.

CHANGES
    Minor.
    Code for writing DIMENSIONS updated to account for bug when
    reading file with QuickView. Renamed xl_write methods to write.
    

VERSION
    This document refers to version 0.09 of Spreadsheet::WriteExcel,
    released Feb 1, 2000.

AUTHOR
    John McNamara (john.exeng@abanet.it)
