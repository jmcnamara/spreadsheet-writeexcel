#!/usr/bin/perl -w

###############################################################################
#
# Example of how to use extend the  Spreadsheet::WriteExcel 7MB limit with
# OLE::Storage_Lite http://search.cpan.org/search?dist=OLE-Storage_Lite
#
# Nov 2000, Kawai, Takanori (Hippo2000)
#   Mail: GCD00051@nifty.ne.jp
#   http://member.nifty.ne.jp/hippo2000
#
# To run this program you need to copy Big.pm and WorkbookBig.pm to
# yourperl/site/lib/Spreadsheet/WriteExcel
#
# Currently the Excel data is transfered to OLE::Storage_Lite as a single
# scalar. This is slow and requires a lot of memory. 
#

# Create a BIG Excel file


use strict;
use Spreadsheet::WriteExcel::Big; # Note the name


my $oExW = Spreadsheet::WriteExcel::Big->new("big.xls");
my $oWorksheet = $oExW->addworksheet();
$oWorksheet->set_column(0, 50, 18);

for(my $iCol=0; $iCol< 50; $iCol++) {
    for(my $iRow=0; $iRow< 6000; $iRow++) {
        $oWorksheet->write($iRow, $iCol, "ROW:$iRow COL:$iCol");
    }
}

$oExW->close();
