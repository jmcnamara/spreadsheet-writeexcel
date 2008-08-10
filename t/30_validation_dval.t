#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Tests for the Excel DVAL structure used in data validation.
#
# reverse('©'), September 2008, John McNamara, jmcnamara@cpan.org
#


use strict;

use Spreadsheet::WriteExcel;
use Test::More tests => 3;


###############################################################################
#
# Tests setup
#
my $test_file           = "temp_test_file.xls";
my $workbook            = Spreadsheet::WriteExcel->new($test_file);
my $worksheet           = $workbook->add_worksheet();
my $target;
my $result;
my $caption;

my $dv_count;
my $obj_id;



###############################################################################
#
# Test 1.
#

$obj_id     = 1;
$dv_count   = 1;

$caption    = " \tData validation: _store_dval($obj_id, $dv_count)";
$target     = join " ",  qw(
                            B2 01 12 00 04 00 00 00 00 00 00 00 00 00 01 00
                            00 00 01 00 00 00
                           );

$result     = unpack_record($worksheet->_store_dval($obj_id, $dv_count));
is($result, $target, $caption);


###############################################################################
#
# Test 2.
#

$obj_id     = -1;
$dv_count   = 1;

$caption    = " \tData validation: _store_dval($obj_id, $dv_count)";
$target     = join " ",  qw(
                            B2 01 12 00 04 00 00 00 00 00 00 00 00 00 FF FF
                            FF FF 01 00 00 00
                           );

$result     = unpack_record($worksheet->_store_dval($obj_id, $dv_count));
is($result, $target, $caption);


###############################################################################
#
# Test 3.
#

$obj_id     = 1;
$dv_count   = 2;

$caption    = " \tData validation: _store_dval($obj_id, $dv_count)";
$target     = join " ",  qw(
                            B2 01 12 00 04 00 00 00 00 00 00 00 00 00 01 00
                            00 00 02 00 00 00

                           );

$result     = unpack_record($worksheet->_store_dval($obj_id, $dv_count));
is($result, $target, $caption);



###############################################################################
#
# Unpack the binary data into a format suitable for printing in tests.
#
sub unpack_record {
    return join ' ', map {sprintf "%02X", $_} unpack "C*", $_[0];
}


# Cleanup
$workbook->close();
unlink $test_file;


__END__



