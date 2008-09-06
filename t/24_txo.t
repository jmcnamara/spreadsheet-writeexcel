#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Tests for some of the internal method used to write the NOTE record that
# is used in cell comments.
#
# reverse('©'), September 2005, John McNamara, jmcnamara@cpan.org
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

my $string;
my $formats;




###############################################################################
#
# Test 1 TXO.
#

$string     = 'aaa';
$caption    = " \t_store_txo()";
$target     = join " ",  qw(
                            B6 01 12 00 12 02 00 00 00 00 00 00 00 00 03 00
                            10 00 00 00 00 00
                           );


$result     = unpack_record($worksheet->_store_txo(length $string));

is($result, $target, $caption);



###############################################################################
#
# Test 2 First CONTINUE record after TXO.
#

$string     = 'aaa';
$caption    = " \t_store_txo_continue_1()";
$target     = join " ",  qw(
                            3C 00 04 00 00 61 61 61
                           );


$result     = unpack_record($worksheet->_store_txo_continue_1($string));

is($result, $target, $caption);



###############################################################################
#
# Test 3 Second CONTINUE record after TXO.
#

$string     = 'aaa';
$caption    = " \t_store_txo_continue_2()";
$target     = join " ",  qw(
                            3C 00 10 00 00 00 00 00 00 00 00 00 03 00 00 00
                            00 00 00 00
                           );

$formats    = [
                [0,               0],
                [length($string), 0],
              ];

$result     = unpack_record($worksheet->_store_txo_continue_2($formats));

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



