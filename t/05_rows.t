#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Check that max/min columns of the Excel ROW record are written correctly.
#
# reverse('©'), October 2007, John McNamara, jmcnamara@cpan.org
#


use strict;

use Spreadsheet::WriteExcel;
use Test::More tests => 7;


###############################################################################
#
# Tests setup
#
my $test_file   = "temp_test_file.xls";
my $workbook    = Spreadsheet::WriteExcel->new($test_file);
my $worksheet;
my @rows;
my @tests;
my $row;
my $col1;
my $col2;

$workbook->compatibility_mode(1);



###############################################################################
#
# Test 1.
#
$row  = 1;
$col1 = 0;
$col2 = 0;
$worksheet = $workbook->add_worksheet();
$worksheet->set_row($row, 15);
push @tests,    [
                    " \tset_row(): row = $row, col1 = $col1, col2 = $col2",
                    {
                        col_min => 0,
                        col_max => 0,
                    }
                ];

###############################################################################
#
# Test 2.
#
$row  = 2;
$col1 = 0;
$col2 = 0;
$worksheet = $workbook->add_worksheet();
$worksheet->write($row, $col1, 'Test');
$worksheet->write($row, $col2, 'Test');
push @tests,    [
                    " \twrite():   row = $row, col1 = $col1, col2 = $col2",
                    {
                        col_min => 0,
                        col_max => 1,
                    }
                ];


###############################################################################
#
# Test 3.
#
$row  = 3;
$col1 = 0;
$col2 = 1;
$worksheet = $workbook->add_worksheet();
$worksheet->write($row, $col1, 'Test');
$worksheet->write($row, $col2, 'Test');
push @tests,    [
                    " \twrite():   row = $row, col1 = $col1, col2 = $col2",
                    {
                        col_min => 0,
                        col_max => 2,
                    }
                ];


###############################################################################
#
# Test 4.
#
$row  = 4;
$col1 = 1;
$col2 = 1;
$worksheet = $workbook->add_worksheet();
$worksheet->write($row, $col1, 'Test');
$worksheet->write($row, $col2, 'Test');
push @tests,    [
                    " \twrite():   row = $row, col1 = $col1, col2 = $col2",
                    {
                        col_min => 1,
                        col_max => 2,
                    }
                ];


###############################################################################
#
# Test 5.
#
$row  = 5;
$col1 = 1;
$col2 = 255;
$worksheet = $workbook->add_worksheet();
$worksheet->write($row, $col1, 'Test');
$worksheet->write($row, $col2, 'Test');
push @tests,    [
                    " \twrite():   row = $row, col1 = $col1, col2 = $col2",
                    {
                        col_min => 1,
                        col_max => 256,
                    }
                ];


###############################################################################
#
# Test 6.
#
$row  = 6;
$col1 = 255;
$col2 = 255;
$worksheet = $workbook->add_worksheet();
$worksheet->write($row, $col1, 'Test');
$worksheet->write($row, $col2, 'Test');
push @tests,    [
                    " \twrite():   row = $row, col1 = $col1, col2 = $col2",
                    {
                        col_min => 255,
                        col_max => 256,
                    }
                ];

###############################################################################
#
# Test 7.
#
$row  = 7;
$col1 = 2;
$col2 = 9;
$worksheet = $workbook->add_worksheet();
$worksheet->set_row($row, 15);
$worksheet->write($row, $col1, 'Test');
$worksheet->write($row, $col2, 'Test');
push @tests,    [
                    " \tset_row + write():   row = $row, col1 = $col1, col2 = $col2",
                    {
                        col_min => 2,
                        col_max => 10,
                    }
                ];




# Read in the row records
$workbook->{_biff_only} = 1;
$workbook->close();

open    XLSFILE, $test_file or die "Couldn't open test file\n";
binmode XLSFILE;

my $header;
my $data;
while (read XLSFILE, $header, 4) {

    my ($record, $length) = unpack 'vv', $header;
    read XLSFILE, $data, $length;

    # Read the row records only.
    next unless $record == 0x0208;
    my ($col_min, $col_max) = unpack 'x2 vv', $data;

    push @rows,
                {
                    col_min => $col_min,
                    col_max => $col_max,
                };
}


for my $i (0 .. @tests -1) {

    is_deeply($rows[$i], $tests[$i]->[1], $tests[$i]->[0]);
}





# Clean up.
unlink $test_file;

__END__



