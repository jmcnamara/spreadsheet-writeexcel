#!/usr/bin/perl -w

###############################################################################
#
# Example of how to use the merge_cells() workbook method with 
# Spreadsheet::WriteExcel.
#
# The usual way to merge cells with Spreadsheet::WriteExcel is to set the merge
# property of a format and to apply that format to the cells to be merged.
# However, in some circumstances this isn't sufficient, for example if the
# merged cells contain a hyperlink or if you wish to merge cells in more than
# one row. In this case you need to cell the merge_cells() workbook method in
# addition to setting the format.
#
# Merged hyperlinks can also be created using write_url_range(). 
#
# reverse('©'), March 2002, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new workbook called simple.xls and add a worksheet
my $workbook  = Spreadsheet::WriteExcel->new("merge3.xls");
my $worksheet = $workbook->addworksheet();



# Create a format that looks like a hyperlink and that has the merge property
# set. Note the border is applied around the merged cells and not around each
# individual cell.
#
my $format = $workbook->addformat(
                                    merge       => 1,
                                    border      => 1,
                                    color       => 'blue',
                                    underline   => 1,
                                 );

#
# Merge cells containing a hyperlink, Method 1: using write_url_range().
#

# Write the cells to be merged
$worksheet->write_url_range("B2:D2", "http://www.perl.com", $format);
$worksheet->write("C2", "", $format);
$worksheet->write("D2", "", $format);

#
# Merge cells containing a hyperlink, Method 2: using merge_cells().
# Note: Call merge_cells *before* writing the cells to be merged.
#
$worksheet->merge_cells("B4:D4");


# Write the cells to be merged
$worksheet->write("B4", "http://www.perl.com", $format);
$worksheet->write("C4", "", $format);
$worksheet->write("D4", "", $format);



#
# Merge cells over two rows using merge_cells
#
$worksheet->merge_cells("B7:D8");

$worksheet->write("B7", "http://www.perl.com", $format);
$worksheet->write("C7", "", $format);
$worksheet->write("D7", "", $format);
$worksheet->write("B8", "", $format);
$worksheet->write("C8", "", $format);
$worksheet->write("D8", "", $format);

