NAME
    Spreadsheet::WriteExcel - Write text and numbers to minimal
    Excel binary file.

DESCRIPTION
    This module can be used to write numbers and text in the native
    Excel binary file format. This is a minimal implementation of an
    Excel file; no formatting can be applied to cells and only a
    single worksheet can be written to a workbook.
    
    This module cannot be used to read an Excel file. Look at the main
    documentation for some suggestions.

    It is intended to be cross-platform, however, this is not
    guaranteed. See the section on portability in the main
    documentation.

INSTALLATION
    Use the standard installation procedure:
        perl Makefile.PL
        make
        make test
        make install

VERSION
    This document refers to version 0.10 of Spreadsheet::WriteExcel,
    released May 13, 2000.

AUTHOR
    John McNamara (writeexcel@eircom.net)
