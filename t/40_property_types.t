#!/usr/bin/perl -w

###############################################################################
#
# Testcases for Spreadsheet::WriteExcel.
#
# Tests for the basic property types used in OLE property sets.
#
# reverse('©'), Auguest 2008, John McNamara, jmcnamara@cpan.org
#


use strict;
use Carp;

use Spreadsheet::WriteExcel::Properties ':testing';
use Time::Local 'timegm';
use Test::More tests => 13;


###############################################################################
#
# Tests setup
#
my $target;
my $result;
my $caption;
my $string;
my $codepage;
my $smiley = chr 0x263A;
my $filetime;


###############################################################################
#
# Test 1. Pack a VT_I2.
#
$caption    = " \tDoc properties: _pack_VT_I2(1252)";
$target     = join " ",  qw(
                            02 00 00 00 E4 04 00 00
                           );

$result     = unpack_record( _pack_VT_I2(1252) );
is($result, $target, $caption);


###############################################################################
#
# Test 2. Pack a VT_LPSTR string and check for padding.
#
$string     = '';
$codepage   = 0x04E4;
$caption    = " \tDoc properties: _pack_VT_LPSTR('$string',\t$codepage')";
$target     = join " ",  qw(
                            1E 00 00 00 01 00 00 00 00 00 00 00
                           );

$result     = unpack_record( _pack_VT_LPSTR($string, $codepage) );
is($result, $target, $caption);


###############################################################################
#
# Test 3. Pack a VT_LPSTR string and check for padding.
#
$string     = 'a';
$codepage   = 0x04E4;
$caption    = " \tDoc properties: _pack_VT_LPSTR('$string',\t$codepage')";
$target     = join " ",  qw(
                            1E 00 00 00 02 00 00 00 61 00 00 00
                           );

$result     = unpack_record( _pack_VT_LPSTR($string, $codepage) );
is($result, $target, $caption);


###############################################################################
#
# Test 4. Pack a VT_LPSTR string and check for padding.
#
$string     = 'bb';
$codepage   = 0x04E4;
$caption    = " \tDoc properties: _pack_VT_LPSTR('$string',\t$codepage')";
$target     = join " ",  qw(
                            1E 00 00 00 03 00 00 00 62 62 00 00
                           );

$result     = unpack_record( _pack_VT_LPSTR($string, $codepage) );
is($result, $target, $caption);


###############################################################################
#
# Test 5. Pack a VT_LPSTR string and check for padding.
#
$string     = 'ccc';
$codepage   = 0x04E4;
$caption    = " \tDoc properties: _pack_VT_LPSTR('$string',\t$codepage')";
$target     = join " ",  qw(
                            1E 00 00 00 04 00 00 00 63 63 63 00
                           );

$result     = unpack_record( _pack_VT_LPSTR($string, $codepage) );
is($result, $target, $caption);


###############################################################################
#
# Test 6. Pack a VT_LPSTR string and check for padding.
#
$string     = 'dddd';
$codepage   = 0x04E4;
$caption    = " \tDoc properties: _pack_VT_LPSTR('$string',\t$codepage')";
$target     = join " ",  qw(
                            1E 00 00 00 05 00 00 00 64 64 64 64 00 00 00 00
                           );

$result     = unpack_record( _pack_VT_LPSTR($string, $codepage) );
is($result, $target, $caption);


###############################################################################
#
# Test 7. Pack a VT_LPSTR string and check for padding.
#
$string     = 'Username';
$codepage   = 0x04E4;
$caption    = " \tDoc properties: _pack_VT_LPSTR('$string',\t$codepage')";
$target     = join " ",  qw(
                            1E 00 00 00 09 00 00 00 55 73 65 72 6E 61 6D 65
                            00 00 00 00
                           );

$result     = unpack_record( _pack_VT_LPSTR($string, $codepage) );
is($result, $target, $caption);


###############################################################################
#
# Test 8. Pack a VT_LPSTR UTF8 string.
#
SKIP: {

skip " \t_pack_VT_LPSTR(utf8). Test requires Perl 5.8 Unicode support.", 1
     if $] < 5.008;

$string     = "$smiley";
$codepage   = 0xFDE9;
$caption    = " \tDoc properties: _pack_VT_LPSTR('\$smiley',\t$codepage')";
$target     = join " ",  qw(
                            1E 00 00 00 04 00 00 00 E2 98 BA 00
                           );

$result     = unpack_record( _pack_VT_LPSTR($string, $codepage) );
is($result, $target, $caption);

}


###############################################################################
#
# Test 9. Pack a VT_LPSTR UTF8 string.
#
SKIP: {

skip " \t_pack_VT_LPSTR(utf8). Test requires Perl 5.8 Unicode support.", 1
     if $] < 5.008;

$string     = "a$smiley";
$codepage   = 0xFDE9;
$caption    = " \tDoc properties: _pack_VT_LPSTR('a\$smiley',\t$codepage')";
$target     = join " ",  qw(
                            1E 00 00 00 05 00 00 00 61 E2 98 BA 00 00 00 00
                           );

$result     = unpack_record( _pack_VT_LPSTR($string, $codepage) );
is($result, $target, $caption);

}


###############################################################################
#
# Test 10. Pack a VT_LPSTR UTF8 string.
#
SKIP: {

skip " \t_pack_VT_LPSTR(utf8). Test requires Perl 5.8 Unicode support.", 1
     if $] < 5.008;

$string     = "aa$smiley";
$codepage   = 0xFDE9;
$caption    = " \tDoc properties: _pack_VT_LPSTR('aa\$smiley',\t$codepage')";
$target     = join " ",  qw(
                            1E 00 00 00 06 00 00 00 61 61 E2 98 BA 00 00 00
                           );

$result     = unpack_record( _pack_VT_LPSTR($string, $codepage) );
is($result, $target, $caption);

}


###############################################################################
#
# Test 11. Pack a VT_LPSTR UTF8 string.
#
SKIP: {

skip " \t_pack_VT_LPSTR(utf8). Test requires Perl 5.8 Unicode support.", 1
     if $] < 5.008;

$string     = "aaa$smiley";
$codepage   = 0xFDE9;
$caption    = " \tDoc properties: _pack_VT_LPSTR('aaa\$smiley',\t$codepage')";
$target     = join " ",  qw(
                            1E 00 00 00 07 00 00 00 61 61 61 E2 98 BA 00 00
                           );

$result     = unpack_record( _pack_VT_LPSTR($string, $codepage) );
is($result, $target, $caption);

}


###############################################################################
#
# Test 12. Pack a VT_LPSTR UTF8 string.
#
SKIP: {

skip " \t_pack_VT_LPSTR(utf8). Test requires Perl 5.8 Unicode support.", 1
     if $] < 5.008;

$string     = "aaaa$smiley";
$codepage   = 0xFDE9;
$caption    = " \tDoc properties: _pack_VT_LPSTR('aaaa\$smiley',\t$codepage')";
$target     = join " ",  qw(
                            1E 00 00 00 08 00 00 00 61 61 61 61 E2 98 BA 00
                           );

$result     = unpack_record( _pack_VT_LPSTR($string, $codepage) );
is($result, $target, $caption);

}


###############################################################################
#
# Test 13. Pack a VT_FILETIME.
#

# Wed Aug 13 01:40:00 2008
# $sec,$min,$hour,$mday,$mon,$year
# We normalise the time using timegm() so that the tests don't fail due to
# different timezones.
$filetime   = [localtime(timegm(0, 40, 0, 13, 7, 108))];

$caption    = " \tDoc properties: _pack_VT_FILETIME()";
$target     = join " ",  qw(
                            40 00 00 00 00 70 EB 1D DD FC C8 01
                           );

$result     = unpack_record( _pack_VT_FILETIME($filetime) );
is($result, $target, $caption);


###############################################################################
#
# Unpack the binary data into a format suitable for printing in tests.
#
sub unpack_record {
    return join ' ', map {sprintf "%02X", $_} unpack "C*", $_[0];
}


__END__
