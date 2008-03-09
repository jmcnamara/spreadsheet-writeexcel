#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Tests to ensure Spreadsheet::WriteExcel::Format->new() and set_properties
# don't do unsafe things.
#


use strict;

use Spreadsheet::WriteExcel::Format;
use Test::More tests => 2;

{
  my $ok = 0;
  eval {
    my $foo = Spreadsheet::WriteExcel::Format->new(0, font => q#'); die ('Error# );
    $ok = 1;
  };
  ok($ok, " Spreadsheet::WriteExcel::Format->new()");
}

{
  my $ok = 0;
  my $format = Spreadsheet::WriteExcel::Format->new();
  eval {
    my $foo = $format->set_properties(size => q#'); die ('Error# );
    $ok = 1;
  };
  ok($ok, " Spreadsheet::WriteExcel::Format->set_properties()");
}
