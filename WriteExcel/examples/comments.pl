#!/usr/bin/perl -w

###############################################################################
#
# This example demonstrates writing cell comments. A cell comment is indicated
# in Excel by a small red triangle in the upper right-hand corner of the cell.
#
# reverse('©'), April 2003, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook  = Spreadsheet::WriteExcel->new("comments.xls");
my $worksheet = $workbook->add_worksheet();

$worksheet->write  (2, 2, "Hello");
$worksheet->write_comment(2, 2,"This is a comment.");

