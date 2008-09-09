#!/usr/bin/perl -w

###############################################################################
#
# Testcases for Spreadsheet::WriteExcel.
#
# Tests for OLE property sets.
#
# reverse('©'), Auguest 2008, John McNamara, jmcnamara@cpan.org
#


use strict;
use Carp;

use Spreadsheet::WriteExcel::Properties ':testing';
use Time::Local 'timegm';
use Test::More tests => 8;


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
my @properties;


###############################################################################
#
# Test 1. Codepage only.
#

@properties =  ([0x0001, 'VT_I2', 0x04E4 ]);
$caption    = " \tDoc properties: create_summary_property_set('Code page')";
$target     = join " ",  qw(
                            FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00
                            00 00 00 00 00 00 00 00 01 00 00 00 E0 85 9F F2
                            F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00
                            18 00 00 00 01 00 00 00 01 00 00 00 10 00 00 00
                            02 00 00 00 E4 04 00 00
                           );

$result     = unpack_record( create_summary_property_set(\@properties) );
is($result, $target, $caption);


###############################################################################
#
# Test 2. Same as previous + Title.
#

@properties =  (
                [0x0001, 'VT_I2',    0x04E4 ],
                [0x0002, 'VT_LPSTR', 'Title'],
               );
$caption    = " \tDoc properties: create_summary_property_set('+ Title')";
$target     = join " ",  qw(
                            FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00
                            00 00 00 00 00 00 00 00 01 00 00 00 E0 85 9F F2
                            F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00
                            30 00 00 00 02 00 00 00 01 00 00 00 18 00 00 00
                            02 00 00 00 20 00 00 00 02 00 00 00 E4 04 00 00
                            1E 00 00 00 06 00 00 00 54 69 74 6C 65 00 00 00
                           );

$result     = unpack_record( create_summary_property_set(\@properties) );
is($result, $target, $caption);


###############################################################################
#
# Test 3. Same as previous + Subject.
#

@properties =  (
                [0x0001, 'VT_I2',    0x04E4   ],
                [0x0002, 'VT_LPSTR', 'Title'  ],
                [0x0003, 'VT_LPSTR', 'Subject'],
               );
$caption    = " \tDoc properties: create_summary_property_set('+ Subject')";
$target     = join " ",  qw(
                            FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00
                            00 00 00 00 00 00 00 00 01 00 00 00 E0 85 9F F2
                            F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00
                            48 00 00 00 03 00 00 00 01 00 00 00 20 00 00 00
                            02 00 00 00 28 00 00 00 03 00 00 00 38 00 00 00
                            02 00 00 00 E4 04 00 00 1E 00 00 00 06 00 00 00
                            54 69 74 6C 65 00 00 00 1E 00 00 00 08 00 00 00
                            53 75 62 6A 65 63 74 00
                           );

$result     = unpack_record( create_summary_property_set(\@properties) );
is($result, $target, $caption);


###############################################################################
#
# Test 4. Same as previous + Author.
#

@properties =  (
                [0x0001, 'VT_I2',    0x04E4   ],
                [0x0002, 'VT_LPSTR', 'Title'  ],
                [0x0003, 'VT_LPSTR', 'Subject'],
                [0x0004, 'VT_LPSTR', 'Author' ],
               );
$caption    = " \tDoc properties: create_summary_property_set('+ Author')";
$target     = join " ",  qw(
                            FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00
                            00 00 00 00 00 00 00 00 01 00 00 00 E0 85 9F F2
                            F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00
                            60 00 00 00 04 00 00 00 01 00 00 00 28 00 00 00
                            02 00 00 00 30 00 00 00 03 00 00 00 40 00 00 00
                            04 00 00 00 50 00 00 00 02 00 00 00 E4 04 00 00
                            1E 00 00 00 06 00 00 00 54 69 74 6C 65 00 00 00
                            1E 00 00 00 08 00 00 00 53 75 62 6A 65 63 74 00
                            1E 00 00 00 07 00 00 00 41 75 74 68 6F 72 00 00
                           );

$result     = unpack_record( create_summary_property_set(\@properties) );
is($result, $target, $caption);


###############################################################################
#
# Test 5. Same as previous + Keywords.
#

@properties =  (
                [0x0001, 'VT_I2',    0x04E4    ],
                [0x0002, 'VT_LPSTR', 'Title'   ],
                [0x0003, 'VT_LPSTR', 'Subject' ],
                [0x0004, 'VT_LPSTR', 'Author'  ],
                [0x0005, 'VT_LPSTR', 'Keywords'],
               );
$caption    = " \tDoc properties: create_summary_property_set('+ Keywords')";
$target     = join " ",  qw(
                            FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00
                            00 00 00 00 00 00 00 00 01 00 00 00 E0 85 9F F2
                            F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00
                            7C 00 00 00 05 00 00 00 01 00 00 00 30 00 00 00
                            02 00 00 00 38 00 00 00 03 00 00 00 48 00 00 00
                            04 00 00 00 58 00 00 00 05 00 00 00 68 00 00 00
                            02 00 00 00 E4 04 00 00 1E 00 00 00 06 00 00 00
                            54 69 74 6C 65 00 00 00 1E 00 00 00 08 00 00 00
                            53 75 62 6A 65 63 74 00 1E 00 00 00 07 00 00 00
                            41 75 74 68 6F 72 00 00 1E 00 00 00 09 00 00 00
                            4B 65 79 77 6F 72 64 73 00 00 00 00
                           );

$result     = unpack_record( create_summary_property_set(\@properties) );
is($result, $target, $caption);


###############################################################################
#
# Test 6. Same as previous + Comments.
#

@properties =  (
                [0x0001, 'VT_I2',    0x04E4    ],
                [0x0002, 'VT_LPSTR', 'Title'   ],
                [0x0003, 'VT_LPSTR', 'Subject' ],
                [0x0004, 'VT_LPSTR', 'Author'  ],
                [0x0005, 'VT_LPSTR', 'Keywords'],
                [0x0006, 'VT_LPSTR', 'Comments'],
               );
$caption    = " \tDoc properties: create_summary_property_set('+ Comments')";
$target     = join " ",  qw(
                            FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00
                            00 00 00 00 00 00 00 00 01 00 00 00 E0 85 9F F2
                            F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00
                            98 00 00 00 06 00 00 00 01 00 00 00 38 00 00 00
                            02 00 00 00 40 00 00 00 03 00 00 00 50 00 00 00
                            04 00 00 00 60 00 00 00 05 00 00 00 70 00 00 00
                            06 00 00 00 84 00 00 00 02 00 00 00 E4 04 00 00
                            1E 00 00 00 06 00 00 00 54 69 74 6C 65 00 00 00
                            1E 00 00 00 08 00 00 00 53 75 62 6A 65 63 74 00
                            1E 00 00 00 07 00 00 00 41 75 74 68 6F 72 00 00
                            1E 00 00 00 09 00 00 00 4B 65 79 77 6F 72 64 73
                            00 00 00 00 1E 00 00 00 09 00 00 00 43 6F 6D 6D
                            65 6E 74 73 00 00 00 00
                           );

$result     = unpack_record( create_summary_property_set(\@properties) );
is($result, $target, $caption);


###############################################################################
#
# Test 7. Same as previous + Last author.
#

@properties =  (
                [0x0001, 'VT_I2',    0x04E4    ],
                [0x0002, 'VT_LPSTR', 'Title'   ],
                [0x0003, 'VT_LPSTR', 'Subject' ],
                [0x0004, 'VT_LPSTR', 'Author'  ],
                [0x0005, 'VT_LPSTR', 'Keywords'],
                [0x0006, 'VT_LPSTR', 'Comments'],
                [0x0008, 'VT_LPSTR', 'Username'],
               );
$caption    = " \tDoc properties: create_summary_property_set('+ Last author')";
$target     = join " ",  qw(
                            FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00
                            00 00 00 00 00 00 00 00 01 00 00 00 E0 85 9F F2
                            F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00
                            B4 00 00 00 07 00 00 00 01 00 00 00 40 00 00 00
                            02 00 00 00 48 00 00 00 03 00 00 00 58 00 00 00
                            04 00 00 00 68 00 00 00 05 00 00 00 78 00 00 00
                            06 00 00 00 8C 00 00 00 08 00 00 00 A0 00 00 00
                            02 00 00 00 E4 04 00 00 1E 00 00 00 06 00 00 00
                            54 69 74 6C 65 00 00 00 1E 00 00 00 08 00 00 00
                            53 75 62 6A 65 63 74 00 1E 00 00 00 07 00 00 00
                            41 75 74 68 6F 72 00 00 1E 00 00 00 09 00 00 00
                            4B 65 79 77 6F 72 64 73 00 00 00 00 1E 00 00 00
                            09 00 00 00 43 6F 6D 6D 65 6E 74 73 00 00 00 00
                            1E 00 00 00 09 00 00 00 55 73 65 72 6E 61 6D 65
                            00 00 00 00
                           );

$result     = unpack_record( create_summary_property_set(\@properties) );
is($result, $target, $caption);


###############################################################################
#
# Test 8. Same as previous + Creation date.
#

# Wed Aug 20 00:20:13 2008
# $sec,$min,$hour,$mday,$mon,$year
# We normalise the time using timegm() so that the tests don't fail due to
# different timezones.
$filetime   = [localtime(timegm(13, 20, 23, 19, 7, 108))];

@properties =  (
                [0x0001, 'VT_I2',       0x04E4    ],
                [0x0002, 'VT_LPSTR',    'Title'   ],
                [0x0003, 'VT_LPSTR',    'Subject' ],
                [0x0004, 'VT_LPSTR',    'Author'  ],
                [0x0005, 'VT_LPSTR',    'Keywords'],
                [0x0006, 'VT_LPSTR',    'Comments'],
                [0x0008, 'VT_LPSTR',    'Username'],
                [0x000C, 'VT_FILETIME', $filetime ],
               );
$caption    = " \tDoc properties: create_summary_property_set('+ Creation date')";
$target     = join " ",  qw(
                            FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00
                            00 00 00 00 00 00 00 00 01 00 00 00 E0 85 9F F2
                            F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00
                            C8 00 00 00 08 00 00 00 01 00 00 00 48 00 00 00
                            02 00 00 00 50 00 00 00 03 00 00 00 60 00 00 00
                            04 00 00 00 70 00 00 00 05 00 00 00 80 00 00 00
                            06 00 00 00 94 00 00 00 08 00 00 00 A8 00 00 00
                            0C 00 00 00 BC 00 00 00 02 00 00 00 E4 04 00 00
                            1E 00 00 00 06 00 00 00 54 69 74 6C 65 00 00 00
                            1E 00 00 00 08 00 00 00 53 75 62 6A 65 63 74 00
                            1E 00 00 00 07 00 00 00 41 75 74 68 6F 72 00 00
                            1E 00 00 00 09 00 00 00 4B 65 79 77 6F 72 64 73
                            00 00 00 00 1E 00 00 00 09 00 00 00 43 6F 6D 6D
                            65 6E 74 73 00 00 00 00 1E 00 00 00 09 00 00 00
                            55 73 65 72 6E 61 6D 65 00 00 00 00 40 00 00 00
                            80 74 89 21 52 02 C9 01
                          );

$result     = unpack_record( create_summary_property_set(\@properties) );
is($result, $target, $caption);


###############################################################################
#
# Unpack the binary data into a format suitable for printing in tests.
#
sub unpack_record {
    return join ' ', map {sprintf "%02X", $_} unpack "C*", $_[0];
}


__END__
