#!/usr/bin/perl -w

###############################################################################
#
# Example of how to extend the Spreadsheet::WriteExcel 7MB limit with
# OLE::Storage_Lite: http://search.cpan.org/search?dist=OLE-Storage_Lite
#
# Nov 2000, Kawai, Takanori (Hippo2000)
#   Mail: GCD00051@nifty.ne.jp
#   http://member.nifty.ne.jp/hippo2000
#


use strict;
use Spreadsheet::WriteExcel::Big; # Note the name


my $workbook  = Spreadsheet::WriteExcel::Big->new("big.xls");
my $worksheet = $workbook->add_worksheet();

$worksheet->set_column(0, 50, 18);

for my $col (0 .. 50) {
    for my $row (0 .. 6000) {
        $worksheet->write($row, $col, "Row: $row Col: $col");
    }
}

$workbook->close();
