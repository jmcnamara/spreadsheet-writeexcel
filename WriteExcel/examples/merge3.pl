#!/usr/bin/perl -w

###############################################################################
#
# Example of how to use Spreadsheet::WriteExcel to write a hyperlink in a
# merged cell. There are two options write_url_range() with a standard merge
# format or merge_cells().
#
# reverse('©'), August 2002, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new workbook and add a worksheet
my $workbook  = Spreadsheet::WriteExcel->new('merge3.xls');
my $worksheet = $workbook->addworksheet();


# Increase the cell size of the merged cells to highlight the formatting.
$worksheet->set_row($_, 30) for (1, 3, 6, 7);
$worksheet->set_column('B:D', 20);


###############################################################################
#
# Example 1: Merge cells containing a hyperlink using write_url_range()
# and the standard Excel 5+ merge property.
#
my $format1 = $workbook->addformat(
                                    merge       => 1,
                                    border      => 1,
                                    underline   => 1,
                                    color       => 'blue',
                                 );

# Write the cells to be merged
$worksheet->write_url_range('B2:D2', 'http://www.perl.com', $format1);
$worksheet->write_blank('C2', $format1);
$worksheet->write_blank('D2', $format1);



###############################################################################
#
# Example 2: Merge cells containing a hyperlink using merge_cells().
#
# Note:
#      1. You should call merge_cells() after you write the cells to be merged
#      2. You must specify a format for every cell in the merged region
#      3. The merge property doesn't have to be set when using merge_cells()
#      4. A border is applied around the merged cells and not around each cell
#      5. merge_cells() doesn't work with Excel versions before Excel 97
#
my $format2 = $workbook->addformat(
                                    border      => 1,
                                    underline   => 1,
                                    color       => 'blue',
                                    align       => 'center',
                                    valign      => 'vcenter',
                                  );

# Merge 3 cells
$worksheet->write('B4', 'http://www.perl.com', $format2);
$worksheet->write_blank('C4', $format2);
$worksheet->write_blank('D4', $format2);

$worksheet->merge_cells('B4:D4');


# Merge 3 cells over two rows
$worksheet->write('B7', 'http://www.perl.com', $format2);
$worksheet->write_blank('C7', $format2);
$worksheet->write_blank('D7', $format2);
$worksheet->write_blank('B8', $format2);
$worksheet->write_blank('C8', $format2);
$worksheet->write_blank('D8', $format2);

$worksheet->merge_cells('B7:D8');

