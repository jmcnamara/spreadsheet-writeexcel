#!/usr/bin/perl

######################################################################
#
# This is an example of how to extend Spreadsheet::WriteExcel to allow
# you to write a large amount of data with a single method call.
#
# The new method is called writemany()
#
# Sam Kington, sam@illuminated.co.uk, November 2000
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook=Spreadsheet::WriteExcel->new("writemany.xls");
my $worksheet=$workbook->add_worksheet("Test");

my @contents;

my $bold=$workbook->add_format;
$bold->set_bold(1);
$bold->set_color("red");
$bold->set_font("Arial");
$bold->set_border(6);

$worksheet->set_column('A:P', 15);

foreach my $row (0..5) {
    foreach my $col (0..5) {
        $contents[$row][$col]="Row ".($row+1)." : Col ".("A".."Z")[$col];
        $worksheet->write($row, $col, $contents[$row][$col], $bold);
    }
}
$worksheet->writemany(7, 0, \@contents, {direction => "col"});
$worksheet->writemany(14,0, \@contents, {direction => "row", format => $bold});
$worksheet->writemany(21,0, [5..20],    {direction => "col", format => $bold});

$workbook->close;


package Spreadsheet::WriteExcel::Worksheet;

######################################################################
#
# writemany ($row, $col, $ref, $options)
#
# Starting at row $row and column $col, calls either write_number() or
# write_string() on the contents of the arraryref $ref.
# $ref may be a simple arrayref (one-dimensional) or an arrayref of
# arrayrefs (two-dimensional)
# $row and $column are zero indexed.
# $options is a hashref with the following keys:
#  direction:  either "row", "col" depending on the type of $ref. This
#              governs the direction values are inserted; it defaults
#              to "row"
#              If scalar @$ref==5 and direction eq "row", values will be
#              inserted from ($row, $col) to ($row+4, $col)
#              If scalar @$ref==10, scalar @{$ref->[0]}==4 and
#              $direction eq "col", values will be inserted from
#              ($row, $col) to ($row+3, $col+9)
#  format: a format object (optional)
#
# Returns: array of return values of called subroutines, in the same
# order as $ref

sub writemany {
    my ($self, $row, $col, $ref, $options)=@_;
    # If this is an arrayref, go through it
    if (ref($ref) eq "ARRAY") {
    # Work out the direction we're going
    my $direction=$options->{direction} || "row";
    # Work out the converse direction
    my $otherdirection={row=>"col",
                col=>"row"}->{$direction};
    # Cycle through
    for (@$ref) {
        $self->writemany($row, $col, $_,
                 {direction => $otherdirection,
                  format => $options->{format} || undef});
        $direction eq "row" ? $row++ : $col++;
    }
    } else {
    # It's a simple scalar value (or something that we don't
    # handle), so pass it through to write
    $self->write($row, $col, $ref, $options->{format});
    }
}




