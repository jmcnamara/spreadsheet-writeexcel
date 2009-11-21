#!/usr/bin/perl -w

###############################################################################
#
# This example demonstrates writing cell comments.
#
# A cell comment is indicated in Excel by a small red triangle in the upper
# right-hand corner of the cell.
#
# For more advanced comment options see comments2.pl.
#
# reverse('©'), November 2005, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook  = Spreadsheet::WriteExcel->new("comments1.xls");
my $worksheet = $workbook->add_worksheet();



$worksheet->write        ('A1', 'Hello'            );
$worksheet->write_comment('A1', 'This is a comment');

__END__
