#!/usr/bin/perl -w

########################################################################
#
# Example of formatting using the Spreadsheet::WriteExcel module via
# property hashes.
#
# Setting format properties via hashes of values is useful when you have
# to deal with a large number of similar formats. Consider for example a
# chess board pattern with black squares, white unformatted squares and
# a border.
#
# This relatively simple example requires 14 separate Format
# objects although there are only 5 different properties: black
# background, top border, bottom border, left border and right border.
#
# Using property hashes it is possible to define these 5 sets of
# properties and then add them together to create the 14 Format
# configurations.
#
# reverse('©'), July 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook  = Spreadsheet::WriteExcel->new("chess.xls");
my $worksheet = $workbook->add_worksheet();


# Some row and column formatting
$worksheet->set_column('B:I', 10);

for my $i (1..8) {
    $worksheet->set_row($i, 50);
}


# Define the property hashes
#
my %black = (
                'fg_color'  => 'black',
                'pattern'   => 1,
            );

my %top     = ( 'top'    => 6 );
my %bottom  = ( 'bottom' => 6 );
my %left    = ( 'left'   => 6 );
my %right   = ( 'right'  => 6 );


# Define the formats
#
my $format01 = $workbook->add_format(%top,    %left          );
my $format02 = $workbook->add_format(%top,    %black         );
my $format03 = $workbook->add_format(%top,                   );
my $format04 = $workbook->add_format(%top,    %right, %black );

my $format05 = $workbook->add_format(%left                   );
my $format06 = $workbook->add_format(%black                  );
my $format07 = $workbook->add_format(                        );
my $format08 = $workbook->add_format(%right,  %black         );
my $format09 = $workbook->add_format(%right                  );
my $format10 = $workbook->add_format(%left,   %black         );

my $format11 = $workbook->add_format(%bottom, %left,  %black );
my $format12 = $workbook->add_format(%bottom                 );
my $format13 = $workbook->add_format(%bottom, %black         );
my $format14 = $workbook->add_format(%bottom, %right         );


# Draw the pattern
$worksheet->write('B2', '', $format01);
$worksheet->write('C2', '', $format02);
$worksheet->write('D2', '', $format03);
$worksheet->write('E2', '', $format02);
$worksheet->write('F2', '', $format03);
$worksheet->write('G2', '', $format02);
$worksheet->write('H2', '', $format03);
$worksheet->write('I2', '', $format04);

$worksheet->write('B3', '', $format10);
$worksheet->write('C3', '', $format07);
$worksheet->write('D3', '', $format06);
$worksheet->write('E3', '', $format07);
$worksheet->write('F3', '', $format06);
$worksheet->write('G3', '', $format07);
$worksheet->write('H3', '', $format06);
$worksheet->write('I3', '', $format09);

$worksheet->write('B4', '', $format05);
$worksheet->write('C4', '', $format06);
$worksheet->write('D4', '', $format07);
$worksheet->write('E4', '', $format06);
$worksheet->write('F4', '', $format07);
$worksheet->write('G4', '', $format06);
$worksheet->write('H4', '', $format07);
$worksheet->write('I4', '', $format08);

$worksheet->write('B5', '', $format10);
$worksheet->write('C5', '', $format07);
$worksheet->write('D5', '', $format06);
$worksheet->write('E5', '', $format07);
$worksheet->write('F5', '', $format06);
$worksheet->write('G5', '', $format07);
$worksheet->write('H5', '', $format06);
$worksheet->write('I5', '', $format09);

$worksheet->write('B6', '', $format05);
$worksheet->write('C6', '', $format06);
$worksheet->write('D6', '', $format07);
$worksheet->write('E6', '', $format06);
$worksheet->write('F6', '', $format07);
$worksheet->write('G6', '', $format06);
$worksheet->write('H6', '', $format07);
$worksheet->write('I6', '', $format08);

$worksheet->write('B7', '', $format10);
$worksheet->write('C7', '', $format07);
$worksheet->write('D7', '', $format06);
$worksheet->write('E7', '', $format07);
$worksheet->write('F7', '', $format06);
$worksheet->write('G7', '', $format07);
$worksheet->write('H7', '', $format06);
$worksheet->write('I7', '', $format09);

$worksheet->write('B8', '', $format05);
$worksheet->write('C8', '', $format06);
$worksheet->write('D8', '', $format07);
$worksheet->write('E8', '', $format06);
$worksheet->write('F8', '', $format07);
$worksheet->write('G8', '', $format06);
$worksheet->write('H8', '', $format07);
$worksheet->write('I8', '', $format08);

$worksheet->write('B9', '', $format11);
$worksheet->write('C9', '', $format12);
$worksheet->write('D9', '', $format13);
$worksheet->write('E9', '', $format12);
$worksheet->write('F9', '', $format13);
$worksheet->write('G9', '', $format12);
$worksheet->write('H9', '', $format13);
$worksheet->write('I9', '', $format14);



