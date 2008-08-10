#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Tests for the packed caption/message strings used in the Excel DV structure
# as part of data validation.
#
# reverse('©'), September 2008, John McNamara, jmcnamara@cpan.org
#


use strict;

use Spreadsheet::WriteExcel;
use Test::More tests => 8;


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
my $max_length;



###############################################################################
#
# Test 1 Empty string.
#

$string      = '';
$max_length  = 32;

$caption    = " \tData validation: _pack_dv_string('', $max_length)";
$target     = join " ",  qw(
                            01 00 00 00
                           );

$result     = unpack_record($worksheet->_pack_dv_string($string, $max_length));
is($result, $target, $caption);


###############################################################################
#
# Test 2 undef.
#

$string      = undef;
$max_length  = 32;

$caption    = " \tData validation: _pack_dv_string(undef, $max_length)";
$target     = join " ",  qw(
                            01 00 00 00
                           );

$result     = unpack_record($worksheet->_pack_dv_string($string, $max_length));
is($result, $target, $caption);


###############################################################################
#
# Test 3 Single space.
#

$string      = ' ';
$max_length  = 32;

$caption    = " \tData validation: _pack_dv_string(' ', $max_length)";
$target     = join " ",  qw(
                            01 00 00 20
                           );

$result     = unpack_record($worksheet->_pack_dv_string($string, $max_length));
is($result, $target, $caption);


###############################################################################
#
# Test 4 Single character.
#

$string      = 'A';
$max_length  = 32;

$caption    = " \tData validation: _pack_dv_string('$string', $max_length)";
$target     = join " ",  qw(
                            01 00 00 41
                           );

$result     = unpack_record($worksheet->_pack_dv_string($string, $max_length));
is($result, $target, $caption);


###############################################################################
#
# Test 5 String longer than 32 characters (for dialog captions).
#

$string      = 'This string is longer than 32 characters';
$max_length  = 32;

$caption    = " \tData validation: _pack_dv_string('$string', $max_length)";
$target     = join " ",  qw(
                            20 00 00 54 68 69 73 20
                            73 74 72 69 6E 67 20 69 73 20 6C 6F 6E 67 65 72
                            20 74 68 61 6E 20 33 32 20 63 68
                           );

$result     = unpack_record($worksheet->_pack_dv_string($string, $max_length));
is($result, $target, $caption);


###############################################################################
#
# Test 6 String longer than 32 characters  (for dialog messages)..
#

$string      = 'ABCD' x 64;
$max_length  = 255;

$caption    = " \tData validation: _pack_dv_string('264 char string', $max_length)";
$target     = join " ",  qw(
                            FF 00 00 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43
                           );

$result     = unpack_record($worksheet->_pack_dv_string($string, $max_length));
is($result, $target, $caption);



###############################################################################
#
# Test 7 Unicode string.
#
SKIP: {

skip " \_pack_dv_string(). Test requires Perl 5.8 Unicode support.", 1
     if $] < 5.008;

$string      = chr 0x20Ac; # Euro symbol
$max_length  = 32;

$caption    = " \tData validation: _pack_dv_string(utf8 string, $max_length)";
$target     = join " ",  qw(
                            01 00 01 AC 20
                           );

$result     = unpack_record($worksheet->_pack_dv_string($string, $max_length));
is($result, $target, $caption);

}


###############################################################################
#
# Test 8 Longer unicode string.
#
SKIP: {

skip " \_pack_dv_string(). Test requires Perl 5.8 Unicode support.", 1
     if $] < 5.008;

$string      = chr(0x20Ac) . '2.99 Foo';
$max_length  = 32;

$caption    = " \tData validation: _pack_dv_string(utf8 string, $max_length)";
$target     = join " ",  qw(
                            09 00 01 AC 20 32 00 2E
                            00 39 00 39 00 20 00 46 00 6F 00 6F 00
                           );

$result     = unpack_record($worksheet->_pack_dv_string($string, $max_length));
is($result, $target, $caption);

}


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



