#!/usr/bin/perl -w

######################################################################
#
# Example of how to use the WriteExcel module to write text and
# numbers to an Excel binary file.
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new Excel file called perl.xls
my $excel = Spreadsheet::WriteExcel->new("perl.xls");

# General syntax is sub(row, column, token)
# Row and column are zero indexed 

# Write some numbers
$excel->xl_write_number(0, 2, 3);       # Writes 3
$excel->xl_write_number(1, 2, 3.00000); # Writes 3
$excel->xl_write_number(2, 2, 3.00001); # Writes 3.00001
$excel->xl_write_number(3, 2, 3.14159); # TeX revison no.?

# Write some text
$excel->xl_write_string(0, 0, "Hi Excel!");

# Write numbers or text
$excel->xl_write(0, 4, 207E9);    # writes a number
$excel->xl_write(1, 4, "207E9");  # writes a number
$excel->xl_write(2, 4, "207 E9"); # writes a string

