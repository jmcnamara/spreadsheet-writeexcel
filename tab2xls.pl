#!/usr/bin/perl -w

######################################################################
#
# Example of how to use the WriteExcel module
#
# The following converts a tab separated file into an Excel file
#
# Usage: tab2xls.pl tabfile.txt newfile.xls
#

use strict;
use Spreadsheet::WriteExcel;


# Check for valid number of arguments
if (($#ARGV < 1) || ($#ARGV > 2)) { 
    die("Usage: tab2xls.pl tabfile.txt newfile.xls\n");
};


# Open the file with tab separated variables
open (TABFILE, "$ARGV[0]") or die "$ARGV[0]: $!";


# Create a new Excel file
my $excel = Spreadsheet::WriteExcel->new("$ARGV[1]");


# Row and column are zero indexed
my $row = 0;
my $col;


while (<TABFILE>) {
    chomp;
    # Split on single tab
    my @Fld = split('\t', $_);

    $col = 0;
    foreach my $token (@Fld) {
        # Write number or string as necessary
        $excel->write($row, $col, $token);
        $col++;
    }
    $row++;
}
