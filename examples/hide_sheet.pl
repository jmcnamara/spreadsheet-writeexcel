#!/usr/bin/perl -w

#######################################################################
#
# Example of how to hide a worksheet with Spreadsheet::WriteExcel.
#
# reverse('©'), April 2005, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook   = Spreadsheet::WriteExcel->new('hidden.xls');
my $worksheet1 = $workbook->add_worksheet();
my $worksheet2 = $workbook->add_worksheet();
my $worksheet3 = $workbook->add_worksheet();

# Sheet2 won't be visible until it is unhidden in Excel.
$worksheet2->hide();

$worksheet1->write(0, 0, 'Sheet2 is hidden');
$worksheet2->write(0, 0, 'How did you find me?');
$worksheet3->write(0, 0, 'Sheet2 is hidden');


__END__
