#!/usr/bin/perl -w

###############################################################################
#
# Example of how to use the Spreadsheet::WriteExcel merge_cells() workbook
# method with with complex formatting and rotation.
#
# Note:
#      1. You should call merge_cells() after you write the cells to be merged
#      2. You must specify a format for every cell in the merged region
#      3. The merge property doesn't have to be set when using merge_cells()
#      4. A border is applied around the merged cells and not around each cell
#      5. merge_cells() doesn't work with Excel versions before Excel 97
#
# reverse('©'), August 2002, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new workbook and add a worksheet
my $workbook  = Spreadsheet::WriteExcel->new('merge5.xls');
my $worksheet = $workbook->addworksheet();


# Increase the cell size of the merged cells to highlight the formatting.
$worksheet->set_row($_, 60)         for (3..8);
$worksheet->set_column($_, $_ , 15) for (1,3,5);


###############################################################################
#
# Rotation 1, letters run from top to bottom
#
my $format1 = $workbook->addformat(
                                    border      => 6,
                                    bold        => 1,
                                    color       => 'red',
                                    valign      => 'vcentre',
                                    align       => 'centre',
                                    rotation    => 1,
                                  );


$worksheet->write('B4', 'Rotation 1: Top to bottom', $format1);
$worksheet->write_blank('B5', $format1);
$worksheet->write_blank('B6', $format1);
$worksheet->write_blank('B7', $format1);
$worksheet->write_blank('B8', $format1);
$worksheet->write_blank('B9', $format1);

$worksheet->merge_cells('B4:B9');


###############################################################################
#
# Rotation 2, 90° anticlockwise
#
my $format2 = $workbook->addformat(
                                    border      => 6,
                                    bold        => 1,
                                    color       => 'red',
                                    valign      => 'vcentre',
                                    align       => 'centre',
                                    rotation    => 2,
                                  );


$worksheet->write('D4', 'Rotation 2: 90° anticlockwise', $format2);
$worksheet->write_blank('D5', $format2);
$worksheet->write_blank('D6', $format2);
$worksheet->write_blank('D7', $format2);
$worksheet->write_blank('D8', $format2);
$worksheet->write_blank('D9', $format2);

$worksheet->merge_cells('D4:D9');


###############################################################################
#
# Rotation 3, 90° clockwise
#
my $format3 = $workbook->addformat(
                                    border      => 6,
                                    bold        => 1,
                                    color       => 'red',
                                    valign      => 'vcentre',
                                    align       => 'centre',
                                    rotation    => 3,
                                  );


$worksheet->write('F4', 'Rotation 3: 90° clockwise', $format3);
$worksheet->write_blank('F5', $format3);
$worksheet->write_blank('F6', $format3);
$worksheet->write_blank('F7', $format3);
$worksheet->write_blank('F8', $format3);
$worksheet->write_blank('F9', $format3);

$worksheet->merge_cells('F4:F9');

