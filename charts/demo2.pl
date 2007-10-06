#!/usr/bin/perl -w

###############################################################################
#
# Simple example of how to add an externally created chart to a Spreadsheet::
# WriteExcel file.
#
#
# This example adds a pie chart extracted from the file Chart2.xls as follows:
#
#   perl chartex.pl -c=demo2 Chart2.xls
#
#
# reverse('©'), September 2004, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook  = Spreadsheet::WriteExcel->new("demo2.xls");
my $worksheet = $workbook->add_worksheet();


# Add the chart extracted using the chartex utility
my $chart     = $workbook->add_chart_ext('demo201.bin', 'Chart1');


# Link the chart to the worksheet data using a dummy formula.
$worksheet->store_formula('=Sheet1!A1');


# Add some extra formats to cover formats used in the charts.
my $chart_font_1 = $workbook->add_format(font_only => 1);
my $chart_font_2 = $workbook->add_format(font_only => 1);
my $chart_font_3 = $workbook->add_format(font_only => 1);
my $chart_font_4 = $workbook->add_format(font_only => 1);
my $chart_font_5 = $workbook->add_format(font_only => 1);

# Add all other formats.
my $bold      = $workbook->add_format(bold => 1);


# Adjust column widths and add some headers
$worksheet->set_column('A:B', 20);

$worksheet->write('A1', 'Module',        $bold);
$worksheet->write('B1', 'No. of lines', $bold);


# Add data to range that the chart refers to.
my @data = (

            ['BIFFwriter.pm',   275],
            ['Big.pm',           99],
            ['Chart.pm',        269],
            ['Format.pm',       724],
            ['Formula.pm',     1410],
            ['OLEwriter.pm',    447],
            ['Utility.pm',      884],
            ['Workbook.pm',    1925],
            ['WorkbookBig.pm',  112],
            ['Worksheet.pm',   3945],
        );

$worksheet->write_col('A2', \@data);