#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Check that the Excel DIMENSIONS record is written correctly.
#
# reverse('©'), October 2007, John McNamara, jmcnamara@cpan.org
#


use strict;

use Spreadsheet::WriteExcel;
use Test::More tests => 31;


###############################################################################
#
# Tests setup
#
my $test_file   = 'temp_test_file.xls';
my $workbook    = Spreadsheet::WriteExcel->new($test_file);
my $format      = $workbook->add_format();
my $worksheet;
my @dims        = qw(row_min row_max col_min col_max);
my $data;
my $caption;
my %results;
my %expected;
my $error;
my $smiley = pack "n", 0x263a;


###############################################################################
#
# Test 1.
#
$caption            = " \tNo worksheet cell data.";

$worksheet          = $workbook->add_worksheet();

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (0, 0, 0, 0);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 2.
#
$caption            = " \tData in cell (0,     0).";

$worksheet          = $workbook->add_worksheet();
$worksheet->write(0, 0, 'Test');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (0, 1, 0, 1);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 3.
#
$caption            = " \tData in cell (0,     255).";

$worksheet          = $workbook->add_worksheet();
$worksheet->write(0, 255, 'Test');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (0, 1, 255, 256);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 4.
#
$caption            = " \tData in cell (65535, 0).";

$worksheet          = $workbook->add_worksheet();
$worksheet->write(65535, 0, 'Test');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (65535, 65536, 0, 1);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 5.
#
$caption            = " \tData in cell (65535, 255).";

$worksheet          = $workbook->add_worksheet();
$worksheet->write(65535, 255, 'Test');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (65535, 65536, 255, 256);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 6.
#
$caption            = " \tData in cell (5,     3).";

$worksheet          = $workbook->add_worksheet();
$worksheet->write(5, 3, 'Test');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (5, 6, 3, 4);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 7.
#
$caption            = " \tset_row() for row 4.";

$worksheet          = $workbook->add_worksheet();
$worksheet->set_row(4, 20);

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (4, 5, 0, 0);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 8.
#
$caption            = " \tset_row() for row 4..6.";

$worksheet          = $workbook->add_worksheet();
$worksheet->set_row(4, 20);
$worksheet->set_row(5, 20);
$worksheet->set_row(6, 20);

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (4, 7, 0, 0);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 9.
#
$caption            = " \tset_column() for row 4.";

$worksheet          = $workbook->add_worksheet();
$worksheet->set_column(4, 4, 20);

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (0, 0, 0, 0);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 10.
#
$caption            = " \tset_column() for row 4..6.";

$worksheet          = $workbook->add_worksheet();
$worksheet->set_column(4, 6, 20);

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (0, 0, 0, 0);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 11.
#
$caption            = " \tData in cell (0, 0) and set_row() for row 4.";

$worksheet          = $workbook->add_worksheet();
$worksheet->write(0, 0, 'Test');
$worksheet->set_row(4, 20);

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (0, 5, 0, 1);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 12.
#
$caption            = " \tData in cell (0, 0) and set_row() for row 4. Reverse order";

$worksheet          = $workbook->add_worksheet();
$worksheet->set_row(4, 20);
$worksheet->write(0, 0, 'Test');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (0, 5, 0, 1);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 13.
#
$caption            = " \tData in cell (5, 3) and set_row() for row 4.";

$worksheet          = $workbook->add_worksheet();
$worksheet->write(5, 3, 'Test');
$worksheet->set_row(4, 20);

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (4, 6, 3, 4);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 14.
#
$caption            = " \tComment in cell (5, 3).";

$worksheet          = $workbook->add_worksheet();
$worksheet->write_comment(5, 3, 'Test');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (5, 6, 3, 4);

is_deeply(\%results, \%expected, $caption);



###############################################################################
#
# Test 15 + 16.
#
$caption            = " \tundef value for row";

$worksheet          = $workbook->add_worksheet();

{
    # Ignore undef warning.
    $^W = 0;
    $error = $worksheet->write_string(undef, 1, 'Test');
};

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (0, 0, 0, 0);

is_deeply(\%results, \%expected, $caption);
is       (-2,        $error,     $caption . ' (return value)');


###############################################################################
#
# Test 17 + 18.
#
$caption            = " \tundef value for col";

$worksheet          = $workbook->add_worksheet();
$error = $worksheet->write(1, undef, 'Test');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (0, 0, 0, 0);

is_deeply(\%results, \%expected, $caption);
is       (-2,        $error,     $caption . ' (return value)');


###############################################################################
#
# Test 19.
#
$caption            = " \tData in cell (5, 3) and (10, 1).";

$worksheet          = $workbook->add_worksheet();
$worksheet->write(5,  3, 'Test');
$worksheet->write(10, 1, 'Test');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (5, 11, 1, 4);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 20.
#
$caption            = " \tData in cell (5, 3) and (10, 5).";

$worksheet          = $workbook->add_worksheet();
$worksheet->write(5,  3, 'Test');
$worksheet->write(10, 5, 'Test');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (5, 11, 3, 6);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 21.
#
$caption            = " \twrite_string()";

$worksheet          = $workbook->add_worksheet();
$worksheet->write_string(5, 3, 'Test');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (5, 6, 3, 4);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 22.
#
$caption            = " \twrite_number()";

$worksheet          = $workbook->add_worksheet();
$worksheet->write_number(5, 3, 5);

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (5, 6, 3, 4);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 23.
#
$caption            = " \twrite_url()";

$worksheet          = $workbook->add_worksheet();
$worksheet->write_url(5, 3, 'http://www.perl.com');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (5, 6, 3, 4);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 24.
#
$caption            = " \twrite_formula()";

$worksheet          = $workbook->add_worksheet();
$worksheet->write_formula(5, 3, '= 1 + 2');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (5, 6, 3, 4);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 25.
#
$caption            = " \twrite_string()";

$worksheet          = $workbook->add_worksheet();
$worksheet->write_string(5, 3, 'Test');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (5, 6, 3, 4);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 26.
#
$caption            = " \twrite_blank()";

$worksheet          = $workbook->add_worksheet();
$worksheet->write_string(5, 3, $format);

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (5, 6, 3, 4);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 27.
#
$caption            = " \twrite_blank(). No format";

$worksheet          = $workbook->add_worksheet();
$worksheet->write_string(5, 3);

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (0, 0, 0, 0);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 28.
#
$caption            = " \twrite_utf16be_string()";

$worksheet          = $workbook->add_worksheet();
$worksheet->write_utf16be_string(5, 3, $smiley);

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (5, 6, 3, 4);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 29.
#
$caption            = " \twrite_utf16le_string()";

$worksheet          = $workbook->add_worksheet();
$worksheet->write_utf16le_string(5, 3, $smiley);

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (5, 6, 3, 4);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 30.
#
$caption            = " \trepeat_formula()";

$worksheet          = $workbook->add_worksheet();

my $formula         = $worksheet->store_formula('=A1 * 3 + 50');
$worksheet->repeat_formula(5, 3, $formula, $format, 'A1', 'A2');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (5, 6, 3, 4);

is_deeply(\%results, \%expected, $caption);


###############################################################################
#
# Test 31.
#
$caption            = " \tmerge_range()";

$worksheet          = $workbook->add_worksheet();
$format             = $workbook->add_format();

$worksheet->merge_range('C6:E8', 'Test', $format);

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (5, 8, 2, 5);

is_deeply(\%results, \%expected, $caption);


# Clean up.
$workbook->close();
unlink $test_file;


__END__



