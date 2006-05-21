#!/usr/bin/perl -w

##############################################################################
#
# Demonstrates how to wrap data from one worksheet onto another.
#
# Excel has a row limit of 65536 rows. Sometimes the amount of row data to be
# written to a file is greater than this limit. In this case it is a useful
# technique to wrap the data from one worksheet onto the next so that we get
# something like the following:
#
#   Sheet1  Row     1  -  65536
#   Sheet2  Row 65537  - 131072
#   Sheet3  Row 131073 - ...
#
# In order to achieve this we use a single worksheet reference and
# reinitialise it to point to a new worksheet when required.
#
# reverse('©'), May 2006, John McNamara, jmcnamara@cpan.org
#


use strict;
use Spreadsheet::WriteExcel;

my $workbook  = Spreadsheet::WriteExcel->new('row_wrap.xls');
my $worksheet = $workbook->add_worksheet();


# Worksheet formatting.
$worksheet->set_column('A:A', 20);


# For the sake of this example we will use a small row limit. In order to use
# the entire row range set the $row_limit to 65536.
my $row_limit = 10;
my $row       = 0;

for my $count (1 .. 2 * $row_limit +10) {

    # When we hit the row limit we redirect the output
    # to a new worksheet and reset the row number.
    if ($row == $row_limit) {
        $worksheet = $workbook->add_worksheet();
        $row = 0;

        # Repeat any worksheet formatting.
        $worksheet->set_column('A:A', 20);
    }

    $worksheet->write($row, 0,  "This is row $count");
    $row++;
}

