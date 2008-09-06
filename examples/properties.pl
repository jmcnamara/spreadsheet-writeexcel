#!/usr/bin/perl -w

##############################################################################
#
# An example of adding document properties to a Spreadsheet::WriteExcel file.
#
# reverse('©'), August 2008, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook  = Spreadsheet::WriteExcel->new('properties.xls');
my $worksheet = $workbook->add_worksheet();


$workbook->set_properties(
    title    => 'This is an example spreadsheet',
    subject  => 'With document properties',
    author   => 'John McNamara',
    manager  => 'Dr. Heinz Doofenshmirtz ',
    company  => 'of Wolves',
    category => 'Example spreadsheets',
    keywords => 'Sample, Example, Properties',
    comments => 'Created with Perl and Spreadsheet::WriteExcel',
);


$worksheet->set_column('A:A', 50);
$worksheet->write('A1', 'Select File->Properties to see the file properties');


__END__
