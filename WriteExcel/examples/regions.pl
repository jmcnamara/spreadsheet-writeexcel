#!/usr/bin/perl -w

###############################################################################
#
# Example of how to use the WriteExcel module to write a basic multiple
# worksheet Excel file.
#
# reverse('©'), March 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new Excel workbook
my $workbook = Spreadsheet::WriteExcel->new("regions.xls");

# Add some worksheets
my $north = $workbook->addworksheet("North");
my $south = $workbook->addworksheet("South");
my $east  = $workbook->addworksheet("East");
my $west  = $workbook->addworksheet("West");

# Add a Format
my $format = $workbook->addformat();
$format->set_bold();
$format->set_color('blue');

# Add a caption to each worksheet
foreach my $worksheet (@{$workbook->worksheets()}) {
    $worksheet->write(0, 0, "Sales", $format);
}

# Write some data
$north->write(0, 1, 200000);
$south->write(0, 1, 100000);
$east->write (0, 1, 150000);
$west->write (0, 1, 100000);

# Set the active worksheet
$south->activate();

# Set the width of the first column
$south->set_column(0, 0, 20);

# Set the active cell
$south->set_selection(0, 1);
