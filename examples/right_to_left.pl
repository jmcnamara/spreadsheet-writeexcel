#!/usr/bin/perl -w

#######################################################################
#
# Example of how to change the default worksheet direction from
# left-to-right to right-to-left as required by some eastern verions
# of Excel.
#
# reverse('©'), January 2006, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook   = Spreadsheet::WriteExcel->new("right_to_left.xls");
my $worksheet1 = $workbook->add_worksheet();
my $worksheet2 = $workbook->add_worksheet();

$worksheet2->right_to_left();

$worksheet1->write(0, 0, 'Hello'); #  A1, B1, C1, ...
$worksheet2->write(0, 0, 'Hello'); # ..., C1, B1, A1



