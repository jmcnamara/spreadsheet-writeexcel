#!/usr/bin/perl -w

###############################################################################
#
# This example adds a method for writing cell comments to the Worksheet class.
# A cell comment is indicated in Excel by a small red triangle in the upper
# right-hand corner of the cell. This method is not included in the Worksheet
# class by default because it is not forward compatible with the Excel 97-2000
# formats.
#
# The method is called write_comment($row, $col, $comment). The comment can be
# up to 30831 chars.
#
# Code is appended to the Spreadsheet::WriteExcel::Worksheet module by reusing
# the package name. This serves as an example of how the module can be extended
# without modifying the code directly.
#
# reverse('©'), March 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook  = Spreadsheet::WriteExcel->new("comments.xls");
my $worksheet = $workbook->addworksheet();

$worksheet->write  (2, 2, "Hello");
$worksheet->write_comment(2, 2,"This is a comment.");




###############################################################################
#
# The following will be appended to the Spreadsheet::WriteExcel::Worksheet
# package by reusing the package name.
#
package Spreadsheet::WriteExcel::Worksheet;


###############################################################################
#
# write_comment($row, $col, $comment)
#
# Write a comment to the specified row and column (zero indexed). The maximum
# comment size is 30831 chars. Excel5 probably accepts 32k-1 chars. However, it
# can only display 30831 chars. Excel 7 and 2000 will crash above 32k-1.
#
# In Excel 5 a comment is referred to as a NOTE.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long comment truncated to 30831 chars
#
sub write_comment {

    my $self      = shift;
    if (@_ < 3) { return -1 }

    my $row       = $_[0];
    my $col       = $_[1];
    my $str       = $_[2];
    my $strlen    = length($_[2]);
    my $str_error = 0;
    my $str_max   = 30831;
    my $note_max  = 2048;

    if ($row >= $self->{_xls_rowmax}) { return -2 }
    if ($col >= $self->{_xls_colmax}) { return -2 }
    if ($row <  $self->{_dim_rowmin}) { $self->{_dim_rowmin} = $row }
    if ($row >  $self->{_dim_rowmax}) { $self->{_dim_rowmax} = $row }
    if ($col <  $self->{_dim_colmin}) { $self->{_dim_colmin} = $col }
    if ($col >  $self->{_dim_colmax}) { $self->{_dim_colmax} = $col }

    # String must be <= 30831 chars
    if ($strlen > $str_max) {
        $str       = substr($str, 0, $str_max);
        $strlen    = $str_max;
        $str_error = -3;
    }

    # A comment can be up to 30831 chars broken into segments of 2048 chars.
    # The first NOTE record contains the total string length. Each subsequent
    # NOTE record contains the length of that segment.
    #
    my $comment = substr($str, 0, $note_max, '');
    $self->_store_comment($row, $col, $comment, $strlen); # First NOTE

    # Subsequent NOTE records
    while ($str) {
        $comment = substr($str, 0, $note_max, '');
        $strlen  = length($comment);
        # Row is -1 to indicate a continuation NOTE
        $self->_store_comment(-1, 0, $comment, $strlen);
    }

    return $str_error;
}


###############################################################################
#
# _store_comment
#
# Store the Excel 5 NOTE record. This format is not compatible with the Excel 7
# record.
#
sub _store_comment {

    my $self      = shift;
    if (@_ < 3) { return -1 }

    my $record    = 0x001C;                 # Record identifier
    my $length ;                            # Bytes to follow

    my $row       = $_[0];                  # Zero indexed row
    my $col       = $_[1];                  # Zero indexed column
    my $str       = $_[2];
    my $strlen    = $_[3];

    # The length of the first record is the total length of the NOTE.
    # Therefore, it can be greater than 2048.
    #
    if ($strlen > 2048) {
        $length = 0x06 + 2048;
    }
    else{
        $length = 0x06 + $strlen;
    }


    my $header    = pack("vv",  $record, $length);
    my $data      = pack("vvv", $row, $col, $strlen);

    $self->_append($header, $data, $str);
}
