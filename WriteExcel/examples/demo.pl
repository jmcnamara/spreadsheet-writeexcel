#!/usr/bin/perl -w

#######################################################################
#
# Demo of some of the features of Spreadsheet::WriteExcel.
# Used to create the project screenshot for Freshmeat.
#
#
# reverse('©'), October 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook   = Spreadsheet::WriteExcel->new("demo.xls");
my $worksheet  = $workbook->addworksheet('Demo');
my $worksheet2 = $workbook->addworksheet('Another sheet');
my $worksheet3 = $workbook->addworksheet('And another');



#######################################################################
#
# Write a general heading
#
$worksheet->set_column('A:B', 32);
my $heading  = $workbook->addformat(
                                        bold    => 1,
                                        color   => 'blue',
                                        size    => 18,
                                        merge   => 1,
                                        );

my @headings = ('Features of Spreadsheet::WriteExcel', '');
$worksheet->write_row('A1', \@headings, $heading);


#######################################################################
#
# Some text examples
#
my $text_format  = $workbook->addformat(
                                            bold    => 1,
                                            italic  => 1,
                                            color   => 'red',
                                            size    => 18,
                                            font    =>'Lucida Calligraphy'
                                        );

$worksheet->write('A2', "Text");
$worksheet->write('B2', "Hello Excel");
$worksheet->write('A3', "Formatted text");
$worksheet->write('B3', "Hello Excel", $text_format);

#######################################################################
#
# Some numeric examples
#
my $num1_format  = $workbook->addformat(num_format => '$#,##0.00');
my $num2_format  = $workbook->addformat(num_format => ' d mmmm yyy');


$worksheet->write('A4', "Numbers");
$worksheet->write('B4', 1234.56);
$worksheet->write('A5', "Formatted numbers");
$worksheet->write('B5', 1234.56, $num1_format);
$worksheet->write('A6', "Formatted numbers");
$worksheet->write('B6', 37257, $num2_format);


#######################################################################
#
# Formulae
#
$worksheet->set_selection('B7');
$worksheet->write('A7', 'Formulas and functions, "=SIN(PI()/4)"');
$worksheet->write('B7', '=SIN(PI()/4)');


#######################################################################
#
# Hyperlinks
#
$worksheet->write('A8', "Hyperlinks");
$worksheet->write('B8',  'http://www.perl.com/' );


#######################################################################
#
# Images
#
$worksheet->write('A9', "Images");
$worksheet->insert_bitmap('B9', 'republic.bmp', 16, 8);


#######################################################################
#
# Misc
#
$worksheet->write('A17', "Page/printer setup");
$worksheet->write('A18', "Multiple worksheets");


