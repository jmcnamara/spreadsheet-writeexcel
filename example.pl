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

# Write some text
$excel->write_string(0, 0, "Hi Excel!");

# Write some numbers
$excel->write_number(2, 0, 3);       # Writes 3
$excel->write_number(2, 1, 3.00000); # Writes 3
$excel->write_number(2, 2, 3.00001); # Writes 3.00001
$excel->write_number(2, 3, 3.14159); # TeX revison no.?


# Write numbers or text
$excel->write(4, 0, 207E9);    # writes a number
$excel->write(4, 1, "207E9");  # writes a number
$excel->write(4, 2, "207 E9"); # writes a string

