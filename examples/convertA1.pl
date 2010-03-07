#!/usr/bin/perl -w

###############################################################################
#
# This program contains helper functions to deal with the Excel A1 cell
# reference  notation.
#
# These functions have been superseded by L<Spreadsheet::WriteExcel::Utility>.
#
# reverse('©'), March 2001, John McNamara, jmcnamara@cpan.org
#

use strict;

print "\n";
print "Cell B7   is equivalent to (";
print join " ", cell_to_rowcol('B7');
print ") in row column notation.\n";

print "Cell \$B7  is equivalent to (";
print join " ", cell_to_rowcol('$B7');
print ") in row column notation.\n";

print "Cell B\$7  is equivalent to (";
print join " ", cell_to_rowcol('B$7');
print ") in row column notation.\n";

print "Cell \$B\$7 is equivalent to (";
print join " ", cell_to_rowcol('$B$7');
print ") in row column notation.\n\n";

print "Row and column (1999, 29)       are equivalent to ";
print rowcol_to_cell(1999, 29),   ".\n";

print "Row and column (1999, 29, 0, 1) are equivalent to ";
print rowcol_to_cell(1999, 29, 0, 1),   ".\n\n";

print "The base cell is:     Z7\n";
print "Increment the row:    ", inc_cell_row('Z7'), "\n";
print "Decrement the row:    ", dec_cell_row('Z7'), "\n";
print "Increment the column: ", inc_cell_col('Z7'), "\n";
print "Decrement the column: ", dec_cell_col('Z7'), "\n\n";


###############################################################################
#
# rowcol_to_cell($row, $col, $row_absolute, $col_absolute)
#
# Convert a zero based row and column reference to a A1 reference. For example
# (0, 2) to C1. $row_absolute, $col_absolute are optional. They are boolean
# values used to indicate if the row or column value is absolute, i.e. if it is
# prefixed by a $ sign: eg. (0, 2, 0, 1) converts to $C1.
#
# Returns: a cell reference string.
#
sub rowcol_to_cell {

    my $row     = $_[0];
    my $col     = $_[1];
    my $row_abs = $_[2] || 0;
    my $col_abs = $_[3] || 0;


    if ($row_abs) {
        $row_abs = '$'
    }
    else {
        $row_abs = ''
    }

    if ($col_abs) {
        $col_abs = '$'
    }
    else {
        $col_abs = ''
    }


    my $int  = int ($col / 26);
    my $frac = $col % 26 +1;

    my $chr1 ='';
    my $chr2 ='';


    if ($frac != 0) {
        $chr2 = chr (ord('A') + $frac -1);
    }

    if ($int > 0) {
        $chr1 = chr (ord('A') + $int  -1);
    }

    $row++;     # Zero index to 1-index

    return $col_abs . $chr1 . $chr2 . $row_abs. $row;
}


###############################################################################
#
# cell_to_rowcol($cell_ref)
#
# Convert an Excel cell reference in A1 notation to a zero based row and column
# reference; converts C1 to (0, 2, 0, 0).
#
# Returns: row, column, row_is_absolute, column_is_absolute
#
#
sub cell_to_rowcol {

    my $cell = shift;

    $cell =~ /(\$?)([A-I]?[A-Z])(\$?)(\d+)/;

    my $col_abs = $1 eq "" ? 0 : 1;
    my $col     = $2;
    my $row_abs = $3 eq "" ? 0 : 1;
    my $row     = $4;

    # Convert base26 column string to number
    # All your Base are belong to us.
    my @chars  = split //, $col;
    my $expn   = 0;
    $col       = 0;

    while (@chars) {
        my $char = pop(@chars); # LS char first
        $col += (ord($char) -ord('A') +1) * (26**$expn);
        $expn++;
    }

    # Convert 1-index to zero-index
    $row--;
    $col--;

    return $row, $col, $row_abs, $col_abs;
}


###############################################################################
#
# inc_cell_row($cell_ref)
#
# Increments the row number of an Excel cell reference in A1 notation.
# For example C3 to C4
#
# Returns: a cell reference string.
#
sub inc_cell_row {

    my $cell = shift;
    my ($row, $col, $row_abs, $col_abs) = cell_to_rowcol($cell);

    $row++;

    return rowcol_to_cell($row, $col, $row_abs, $col_abs);
}


###############################################################################
#
# dec_cell_row($cell_ref)
#
# Decrements the row number of an Excel cell reference in A1 notation.
# For example C4 to C3
#
# Returns: a cell reference string.
#
sub dec_cell_row {

    my $cell = shift;
    my ($row, $col, $row_abs, $col_abs) = cell_to_rowcol($cell);

    $row--;

    return rowcol_to_cell($row, $col, $row_abs, $col_abs);
}


###############################################################################
#
# inc_cell_col($cell_ref)
#
# Increments the column number of an Excel cell reference in A1 notation.
# For example C3 to D3
#
# Returns: a cell reference string.
#
sub inc_cell_col {

    my $cell = shift;
    my ($row, $col, $row_abs, $col_abs) = cell_to_rowcol($cell);

    $col++;

    return rowcol_to_cell($row, $col, $row_abs, $col_abs);
}


###############################################################################
#
# dec_cell_col($cell_ref)
#
# Decrements the column number of an Excel cell reference in A1 notation.
# For example D3 to C3
#
# Returns: a cell reference string.
#
sub dec_cell_col {

    my $cell = shift;
    my ($row, $col, $row_abs, $col_abs) = cell_to_rowcol($cell);

    $col--;

    return rowcol_to_cell($row, $col, $row_abs, $col_abs);
}

