#!/usr/bin/perl -w

###############################################################################
#
# Simple example of how to add an externally created chart to a Spreadsheet::
# WriteExcel file.
#
#
# This example adds a line chart extracted from the file Chart1.xls as follows:
#
#   perl chartex.pl -c=demo1 Chart1.xls
#
#
# reverse('©'), September 2004, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook  = Spreadsheet::WriteExcel->new("demo1.xls");
my $worksheet = $workbook->add_worksheet();


# Add the chart extracted using the chartex utility
my $chart     = $workbook->add_chart_ext('demo101.bin', 'Chart1');


# Link the chart to the worksheet data using a dummy formula.
$worksheet->store_formula('=Sheet1!A1');


# Add some extra formats to cover formats used in the charts.
my $chart_font_1 = $workbook->add_format(font_only => 1);
my $chart_font_2 = $workbook->add_format(font_only => 1);
my $chart_font_3 = $workbook->add_format(font_only => 1);

# Add all other formats (if any).

# Add data to range that the chart refers to.
my @nums    = (0, 1, 2, 3, 4,  5,  6,  7,  8,  9,  10 );
my @squares = (0, 1, 4, 9, 16, 25, 36, 49, 64, 81, 100);

$worksheet->write_col('A1', \@nums   );
$worksheet->write_col('B1', \@squares);

