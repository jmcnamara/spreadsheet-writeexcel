#!/usr/bin/perl -w

###############################################################################
#
# Example of formatting using the Spreadsheet::WriteExcel module
#
# This example shows how to merge two or more cells. See also the merge2.pl
# example.
#
# Dec 2000, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new workbook and add a worksheet
my $workbook  = Spreadsheet::WriteExcel->new("merge1.xls");
my $worksheet = $workbook->addworksheet();

# Set the column width for columns 2 and 3
$worksheet->set_column(1, 3, 20);

# Set the row height for row 2
$worksheet->set_row(2, 30);


# Create a border format
my $border1 = $workbook->addformat();
$border1->set_align('merge');


# Only one cell should contain text, the others should be blank.
$worksheet->write      (2, 1, "Merged Cells", $border1);
$worksheet->write_blank(2, 2,                 $border1);
$worksheet->write_blank(2, 3,                 $border1);



