#!/usr/bin/perl -w

###############################################################################
#
# Example of how to use the Spreadsheet::WriteExcel module to send an Excel
# file to a browser in a CGI program.
#
# Dec 2000, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Send the content type
print "Content-type: application/vnd.ms-excel\n\n";


# Create a new workbook and add a worksheet. The special Perl filehandle - will
# redirect the output to STDOUT
#
my $workbook  = Spreadsheet::WriteExcel->new("-");
my $worksheet = $workbook->addworksheet();


# Set the column width for column 1
$worksheet->set_column(0, 0, 20);


# Create a format
my $format = $workbook->addformat();
$format->set_bold();
$format->set_size(15);
$format->set_color('blue');


# Write to the workbook
$workbook->write(0, 0, "Hi Excel!", $format);
