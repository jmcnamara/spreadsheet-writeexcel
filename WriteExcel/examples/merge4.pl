#!/usr/bin/perl -w

###############################################################################
#
# Example of how to use the Spreadsheet::WriteExcel merge_cells() workbook
# method with with complex formatting.
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
my $workbook  = Spreadsheet::WriteExcel->new('merge4.xls');
my $worksheet = $workbook->addworksheet();


# Increase the cell size of the merged cells to highlight the formatting.
$worksheet->set_row($_, 30) for (1..11);
$worksheet->set_column('B:D', 20);


###############################################################################
#
# Example 1: Text centered vertically and horizontally
#
my $format1 = $workbook->addformat(
                                    border  => 6,
                                    bold    => 1,
                                    color   => 'red',
                                    valign  => 'vcenter',
                                    align   => 'center',
                                   );



$worksheet->write('B2', 'Centered vertically and horizontally', $format1);
$worksheet->write_blank('C2', $format1);
$worksheet->write_blank('D2', $format1);
$worksheet->write_blank('B3', $format1);
$worksheet->write_blank('C3', $format1);
$worksheet->write_blank('D3', $format1);

$worksheet->merge_cells('B2:D3');


###############################################################################
#
# Example 2: Text aligned to the top and left
#
my $format2 = $workbook->addformat(
                                    border  => 6,
                                    bold    => 1,
                                    color   => 'red',
                                    valign  => 'top',
                                    align   => 'left',
                                  );



$worksheet->write('B5', 'Aligned to the top and left', $format2);
$worksheet->write_blank('C5', $format2);
$worksheet->write_blank('D5', $format2);
$worksheet->write_blank('B6', $format2);
$worksheet->write_blank('C6', $format2);
$worksheet->write_blank('D6', $format2);

$worksheet->merge_cells('B5:D6');


###############################################################################
#
# Example 3:  Text aligned to the bottom and right
#
my $format3 = $workbook->addformat(
                                    border  => 6,
                                    bold    => 1,
                                    color   => 'red',
                                    valign  => 'bottom',
                                    align   => 'right',
                                  );



$worksheet->write('B8', 'Aligned to the bottom and right', $format3);
$worksheet->write_blank('C8', $format3);
$worksheet->write_blank('D8', $format3);
$worksheet->write_blank('B9', $format3);
$worksheet->write_blank('C9', $format3);
$worksheet->write_blank('D9', $format3);

$worksheet->merge_cells('B8:D9');


###############################################################################
#
# Example 4:  Text justified (i.e. wrapped) in the cell
#
my $format4 = $workbook->addformat(
                                    border  => 6,
                                    bold    => 1,
                                    color   => 'red',
                                    valign  => 'top',
                                    align   => 'justify',
                                  );



$worksheet->write('B11', 'Justified: '.'so on and ' x18, $format4);
$worksheet->write_blank('C11', $format4);
$worksheet->write_blank('D11', $format4);
$worksheet->write_blank('B12', $format4);
$worksheet->write_blank('C12', $format4);
$worksheet->write_blank('D12', $format4);

$worksheet->merge_cells('B11:D12');
