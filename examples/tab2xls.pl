#!/usr/bin/perl -w

###############################################################################
#
# Example of how to use the WriteExcel module
#
# The following converts a tab separated file into an Excel file
#
# Usage: tab2xls.pl tabfile.txt newfile.xls
#
#
# NOTE: This is only a simple conversion utility for illustrative purposes.
# For converting a CSV or Tab separated or any other type of delimited
# text file to Excel I recommend the more rigorous csv2xls program that is
# part of H.Merijn Brand's Text::CSV_XS module distro.
#
# See the examples/csv2xls link here:
#     L<http://search.cpan.org/~hmbrand/Text-CSV_XS/MANIFEST>
#
# reverse('©'), March 2001, John McNamara, jmcnamara@cpan.org
#


use strict;
use Spreadsheet::WriteExcel;


# Check for valid number of arguments
if (($#ARGV < 1) || ($#ARGV > 2)) {
    die("Usage: tab2xls tabfile.txt newfile.xls\n");
};


# Open the tab delimited file
open (TABFILE, $ARGV[0]) or die "$ARGV[0]: $!";


# Create a new Excel workbook
my $workbook  = Spreadsheet::WriteExcel->new($ARGV[1]);
my $worksheet = $workbook->add_worksheet();

# Row and column are zero indexed
my $row = 0;

while (<TABFILE>) {
    chomp;
    # Split on single tab
    my @Fld = split('\t', $_);

    my $col = 0;
    foreach my $token (@Fld) {
        $worksheet->write($row, $col, $token);
        $col++;
    }
    $row++;
}
