#!/usr/bin/perl -w

###############################################################################
#
# Simple example of merging cells using the Spreadsheet::WriteExcel module.
#
# This example shows how to merge two or more cells.
#
# reverse('©'), August 2002, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new workbook and add a worksheet
my $workbook  = Spreadsheet::WriteExcel->new("merge1.xls");
my $worksheet = $workbook->addworksheet();


# Increase the cell size of the merged cells to highlight the formatting.
$worksheet->set_column('B:D', 20);
$worksheet->set_row(2, 30);


# Create a merge format
my $format = $workbook->addformat(merge => 1);


# Only one cell should contain text, the others should be blank.
$worksheet->write      (2, 1, "Merged Cells", $format);
$worksheet->write_blank(2, 2,                 $format);
$worksheet->write_blank(2, 3,                 $format);

