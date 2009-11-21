#!/usr/bin/perl -w

#######################################################################
#
# Example of how to write Spreadsheet::WriteExcel formulas with a user
# specified result.
#
# This is generally only required when writing a spreadsheet for an
# application other than Excel where the formula isn't evaluated.
#
# reverse('©'), August 2005, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook  = Spreadsheet::WriteExcel->new('formula_result.xls');
my $worksheet = $workbook->add_worksheet();
my $format    = $workbook->add_format(color => 'blue');


$worksheet->write('A1', '=1+2');
$worksheet->write('A2', '=1+2',                     $format, 4);
$worksheet->write('A3', '="ABC"',                   undef,   'DEF');
$worksheet->write('A4', '=IF(A1 > 1, TRUE, FALSE)', undef,   'TRUE');
$worksheet->write('A5', '=1/0',                     undef,   '#DIV/0!');


__END__
