#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Check that we calculate the correct bucket size and number for the EXTSST
# record. The data is taken from actual Excel files.
#
# reverse('©'), October 2007, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;
use Test::More tests => 56;


###############################################################################
#
# Tests setup
#
my $test_file   = "temp_test_file.xls";
my $workbook    = Spreadsheet::WriteExcel->new($test_file);


my @tests = (  # Unique     Number of   Bucket
               # strings    buckets       size
               [0,          0,               8],
               [1,          1,               8],
               [7,          1,               8],
               [8,          1,               8],
               [15,         2,               8],
               [16,         2,               8],
               [17,         3,               8],
               [32,         4,               8],
               [33,         5,               8],
               [64,         8,               8],
               [128,        16,              8],
               [256,        32,              8],
               [512,        64,              8],
               [1023,       128,             8],
               [1024,       114,             9],
               [1025,       114,             9],
               [2048,       121,            17],
               [4096,       125,            33],
               [4097,       125,            33],
               [8192,       127,            65],
               [8193,       127,            65],
               [9000,       127,            71],
               [10000,      127,            79],
               [16384,      128,           129],
               [262144,     128,          2049],
               [1048576,    128,          8193],
               [4194304,    128,         32769],
               [8257536,    128,         64513],
            );



for my $test (@tests) {

    my $str_unique = $test->[0];

    $workbook->{_str_unique} = $str_unique ;
    $workbook->_calculate_extsst_size();

    is($workbook->{_extsst_buckets},     $test->[1],
        " \tBucket number for $str_unique  strings");
    is($workbook->{_extsst_bucket_size}, $test->[2],
        " \tBucket size   for $str_unique  strings");
}


# Clean up.
$workbook->{_str_unique} = 0;
$workbook->close();
unlink $test_file;
