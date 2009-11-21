#!/usr/bin/perl -w

###############################################################################
#
# Simple example of merging cells using the Spreadsheet::WriteExcel module
#
# This example merges three cells using the "Centre Across Selection"
# alignment which was the Excel 5 method of achieving a merge. For a more
# modern approach use the merge_range() worksheet method instead.
# See the merge3.pl - merge6.pl programs.
#
# reverse('©'), August 2002, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new workbook and add a worksheet
my $workbook  = Spreadsheet::WriteExcel->new("merge2.xls");
my $worksheet = $workbook->add_worksheet();


# Increase the cell size of the merged cells to highlight the formatting.
$worksheet->set_column(1, 2, 30);
$worksheet->set_row(2, 40);


# Create a merged format
my $format = $workbook->add_format(
                                        center_across   => 1,
                                        bold            => 1,
                                        size            => 15,
                                        pattern         => 1,
                                        border          => 6,
                                        color           => 'white',
                                        fg_color        => 'green',
                                        border_color    => 'yellow',
                                        align           => 'vcenter',
                                  );


# Only one cell should contain text, the others should be blank.
$worksheet->write      (2, 1, "Center across selection", $format);
$worksheet->write_blank(2, 2,                            $format);

