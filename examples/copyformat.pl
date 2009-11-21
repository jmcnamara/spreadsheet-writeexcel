#!/usr/bin/perl -w

###############################################################################
#
# Example of how to use the format copying method with Spreadsheet::WriteExcel.
#
# This feature isn't required very often.
#
# reverse('©'), March 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create workbook1
my $workbook1       = Spreadsheet::WriteExcel->new("workbook1.xls");
my $worksheet1      = $workbook1->add_worksheet();
my $format1a        = $workbook1->add_format();
my $format1b        = $workbook1->add_format();

# Create workbook2
my $workbook2       = Spreadsheet::WriteExcel->new("workbook2.xls");
my $worksheet2      = $workbook2->add_worksheet();
my $format2a        = $workbook2->add_format();
my $format2b        = $workbook2->add_format();


# Create a global format object that isn't tied to a workbook
my $global_format   = Spreadsheet::WriteExcel::Format->new();

# Set the formatting
$global_format->set_color('blue');
$global_format->set_bold();
$global_format->set_italic();

# Create another example format
$format1b->set_color('red');

# Copy the global format properties to the worksheet formats
$format1a->copy($global_format);
$format2a->copy($global_format);

# Copy a format from worksheet1 to worksheet2
$format2b->copy($format1b);

# Write some output
$worksheet1->write(0, 0, "Ciao", $format1a);
$worksheet1->write(1, 0, "Ciao", $format1b);

$worksheet2->write(0, 0, "Hello", $format2a);
$worksheet2->write(1, 0, "Hello", $format2b);

