#!/usr/bin/perl -w

###############################################################################
#
# Example of how to use the WriteExcel module to write text and
# numbers to an Excel binary file.
#
# Dec 2000, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new workbook called simple.xls and add a worksheet
my $workbook  = Spreadsheet::WriteExcel->new("simple.xls");
my $worksheet = $workbook->addworksheet();

# General syntax is sub(row, column, token)
# Row and column are zero indexed

# Write some text
$worksheet->write_string(0, 0, "Hi Excel!");

# Write some numbers
$worksheet->write_number(2, 0, 3);          # Writes 3
$worksheet->write_number(2, 1, 3.00000);    # Writes 3
$worksheet->write_number(2, 2, 3.00001);    # Writes 3.00001
$worksheet->write_number(2, 3, 3.14159);    # TeX revision no.?


# Write numbers or text
$worksheet->write(4, 0, 207E9);             # writes a number
$worksheet->write(4, 1, "207E9");           # writes a number
$worksheet->write(4, 2, "207 E9");          # writes a string

