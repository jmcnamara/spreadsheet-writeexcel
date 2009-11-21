#!/usr/bin/perl -w

###############################################################################
#
# Example of how to create defined names in a Spreadsheet::WriteExcel file.
#
# This method is used to defined a name that can be used to represent a value,
# a single cell or a range of cells in a workbook.
#
# reverse('©'), September 2008, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook   = Spreadsheet::WriteExcel->new('defined_name.xls');
my $worksheet1 = $workbook->add_worksheet();
my $worksheet2 = $workbook->add_worksheet();


$workbook->define_name('Exchange_rate', '=0.96');
$workbook->define_name('Sales',         '=Sheet1!$G$1:$H$10');
$workbook->define_name('Sheet2!Sales',  '=Sheet2!$G$1:$G$10');


for my $worksheet ($workbook->sheets()) {
    $worksheet->set_column('A:A', 45);
    $worksheet->write('A2', 'This worksheet contains some defined names,');
    $worksheet->write('A3', 'See the Insert -> Name -> Define dialog.');

}


$worksheet1->write('A4', '=Exchange_rate');

__END__

