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
# reverse('©'), June 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new workbook called simple.xls and add a worksheet
my $workbook  = Spreadsheet::WriteExcel->new("merge3.xls");
my $worksheet = $workbook->addworksheet();



# Create a format that looks like a hyperlink and that has the merge property
# set.
#
my $format = $workbook->addformat();
$format->set_color('blue');
$format->set_underline();
$format->set_merge();


# Call merge_cells *before* writing the cells to be merged.
#
# Note: merge_cells("B4:D4") is equivalent to merge_cells(3, 3, 1, 3);
#
$worksheet->merge_cells("B4:D4");


# Write the cells to be merged
$worksheet->write("B4", "http://www.perl.com", $format);
$worksheet->write("C4", "", $format);
$worksheet->write("D4", "", $format);


# Merge cells over two rows

$worksheet->merge_cells("B7:D8");

$worksheet->write("B7", "http://www.perl.com", $format);
$worksheet->write("C7", "", $format);
$worksheet->write("D7", "", $format);
$worksheet->write("B8", "", $format);
$worksheet->write("C8", "", $format);
$worksheet->write("D8", "", $format);

