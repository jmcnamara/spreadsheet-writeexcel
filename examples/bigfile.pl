#!/usr/bin/perl -w

###############################################################################
#
# Example of creating a Spreadsheet::WriteExcel that is larger than the
# default 7MB limit.
#
# This is exactly that same as any other Spreadsheet::WriteExcel program except
# that is requires that the OLE::Storage module is installed.
#
# reverse('©'), Jan 2007, John McNamara, jmcnamara@cpan.org


use strict;
use Spreadsheet::WriteExcel;


my $workbook  = Spreadsheet::WriteExcel->new('bigfile.xls');
my $worksheet = $workbook->add_worksheet();

$worksheet->set_column(0, 50, 18);

for my $col (0 .. 50) {
    for my $row (0 .. 6000) {
        $worksheet->write($row, $col, "Row: $row Col: $col");
    }
}

__END__
