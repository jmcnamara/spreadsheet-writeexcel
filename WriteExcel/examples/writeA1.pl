#!/usr/bin/perl -w

###############################################################################
#
# This is an example of how to extend the Spreadsheet::WriteExcel module.
#
# Code is appended to the Spreadsheet::WriteExcel::Worksheet module by reusing
# the package name. The new code provides a write() method that allows you to
# use Excels A1 style cell references.  This is not particularly useful but it
# serves as an example of how the module can be extended without modifying the
# code directly.
#
# reverse('©'), March 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new workbook called simple.xls and add a worksheet
my $workbook  = Spreadsheet::WriteExcel->new("writeA1.xls");
my $worksheet = $workbook->add_worksheet();

# Write numbers or text
$worksheet->write  (0, 0, "Hello");
$worksheet->writeA1("A3", "A3"   );
$worksheet->writeA1("A5", 1.2345 );


###############################################################################
#
# The following will be appended to the Spreadsheet::WriteExcel::Worksheet
# package.
#

package Spreadsheet::WriteExcel::Worksheet;

###############################################################################
#
# writeA1($cell, $token, $format)
#
# Convert $cell from Excel A1 notation to $row, $col notation and
# call write() on $token.
#
# Returns: return value of called subroutine or -4 for invalid cell
# reference.
#
sub writeA1 {
    my $self = shift;
    my $cell = shift;
    my $col;
    my $row;

    if ($cell =~ /([A-z]+)(\d+)/) {
       ($row, $col) = _convertA1($2, $1);
       $self->write($row, $col, @_);
    } else {
        return -4;
    }
}

###############################################################################
#
# _convertA1($row, $col)
#
# Convert Excel A1 notation to $row, $col notation. Convert base26 column
# string to a number.
#
sub _convertA1 {
    my $row    = $_[0];
    my $col    = $_[1]; # String in AA notation

    my @chars  = split //, $col;
    my $expn   = 0;
    $col       = 0;

    while (@chars) {
        my $char = uc(pop(@chars)); # LS char first
        $col += (ord($char) -ord('A') +1) * (26**$expn);
        $expn++;
    }

    # Convert 1 index to 0 index
    $row--;
    $col--;

    return($row, $col);
}
