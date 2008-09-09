#!/usr/bin/perl -w

###############################################################################
#
# Testcases for Spreadsheet::WriteExcel.
#
# Tests for Workbook property_sets() interface.
#
# reverse('©'), Auguest 2008, John McNamara, jmcnamara@cpan.org
#


use strict;
use Carp;

use Spreadsheet::WriteExcel;
use Spreadsheet::WriteExcel::Properties ':testing';
use Time::Local 'timegm';
use Test::More tests => 17;


###############################################################################
#
# Tests setup
#
my $test_file   = "temp_test_file.xls";
my $workbook    = Spreadsheet::WriteExcel->new($test_file);
my $worksheet   = $workbook->add_worksheet();

my $target;
my $result;
my $caption;
my $string;
my $codepage;
my $smiley      = chr 0x263A;
my $filetime;
my @properties;
my %params;
my @strings;


###############################################################################
#
# Test 1. _get_property_set_codepage() for default latin1 strings.
#
%params =   (
                title       => 'Title',
                subject     => 'Subject',
                author      => 'Author',
                keywords    => 'Keywords',
                comments    => 'Comments',
                last_author => 'Username',
            );

@strings = qw(title subject author keywords comments last_author);


$caption    = " \t_get_property_set_codepage('latin1')";
$target     = 0x04E4;

$result     = $workbook->_get_property_set_codepage(\%params, \@strings);
is($result, $target, $caption);


###############################################################################
#
# Test 2. _get_property_set_codepage() for manual utf8 strings.
#

%params =   (
                title       => 'Title',
                subject     => 'Subject',
                author      => 'Author',
                keywords    => 'Keywords',
                comments    => 'Comments',
                last_author => 'Username',
                utf8        => 1,
            );

@strings = qw(title subject author keywords comments last_author);


$caption    = " \t_get_property_set_codepage('utf8')";
$target     = 0xFDE9;

$result     = $workbook->_get_property_set_codepage(\%params, \@strings);
is($result, $target, $caption);


###############################################################################
#
# Test 3. _get_property_set_codepage() for perl 5.8 utf8 strings.
#
SKIP: {

skip " \t_get_property_set_codepage('utf8'). Requires Perl 5.8 Unicode.", 1
     if $] < 5.008;

%params =   (
                title       => 'Title' . $smiley,
                subject     => 'Subject',
                author      => 'Author',
                keywords    => 'Keywords',
                comments    => 'Comments',
                last_author => 'Username',
            );

@strings = qw(title subject author keywords comments last_author);


$caption    = " \t_get_property_set_codepage('utf8')";
$target     = 0xFDE9;

$result     = $workbook->_get_property_set_codepage(\%params, \@strings);
is($result, $target, $caption);
}


###############################################################################
#
# Note, the "created => undef" parameters in some of the following tests is
# used to avoid adding the default date to the property sets.


###############################################################################
#
# Test 4. Codepage only.
#

$workbook->set_properties(
                            created     => undef,
                         );

$caption    = " \tset_properties(codepage)";
$target     = join " ",  qw(
                            FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00
                            00 00 00 00 00 00 00 00 01 00 00 00 E0 85 9F F2
                            F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00
                            18 00 00 00 01 00 00 00 01 00 00 00 10 00 00 00
                            02 00 00 00 E4 04 00 00
                           );

$result     = unpack_record( $workbook->{summary} );
is($result, $target, $caption);


###############################################################################
#
# Test 5. Same as previous + Title.
#

$workbook->set_properties(
                            title       => 'Title',
                            created     => undef,
                         );

$caption    = " \tset_properties('Title')";
$target     = join " ",  qw(
                            FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00
                            00 00 00 00 00 00 00 00 01 00 00 00 E0 85 9F F2
                            F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00
                            30 00 00 00 02 00 00 00 01 00 00 00 18 00 00 00
                            02 00 00 00 20 00 00 00 02 00 00 00 E4 04 00 00
                            1E 00 00 00 06 00 00 00 54 69 74 6C 65 00 00 00
                           );

$result     = unpack_record( $workbook->{summary} );
is($result, $target, $caption);


###############################################################################
#
# Test 6. Same as previous + Subject.
#

$workbook->set_properties(
                            title       => 'Title',
                            subject     => 'Subject',
                            created     => undef,
                         );

$caption    = " \tset_properties('+ Subject')";
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

$result     = unpack_record( $workbook->{summary} );
is($result, $target, $caption);


###############################################################################
#
# Test 7. Same as previous + Author.
#

$workbook->set_properties(
                            title       => 'Title',
                            subject     => 'Subject',
                            author      => 'Author',
                            created     => undef,
                         );

$caption    = " \tset_properties('+ Author')";
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

$result     = unpack_record( $workbook->{summary} );
is($result, $target, $caption);


###############################################################################
#
# Test 8. Same as previous + Keywords.
#

$workbook->set_properties(
                            title       => 'Title',
                            subject     => 'Subject',
                            author      => 'Author',
                            keywords    => 'Keywords',
                            created     => undef,
                         );

$caption    = " \tset_properties('+ Keywords')";
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

$result     = unpack_record( $workbook->{summary} );
is($result, $target, $caption);


###############################################################################
#
# Test 9. Same as previous + Comments.
#

$workbook->set_properties(
                            title       => 'Title',
                            subject     => 'Subject',
                            author      => 'Author',
                            keywords    => 'Keywords',
                            comments    => 'Comments',
                            created     => undef,
                         );

$caption    = " \tset_properties('+ Comments')";
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

$result     = unpack_record( $workbook->{summary} );
is($result, $target, $caption);


###############################################################################
#
# Test 10. Same as previous + Last author.
#

$workbook->set_properties(
                            title       => 'Title',
                            subject     => 'Subject',
                            author      => 'Author',
                            keywords    => 'Keywords',
                            comments    => 'Comments',
                            last_author => 'Username',
                            created     => undef,
                         );

$caption    = " \tset_properties('+ Last author')";
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

$result     = unpack_record( $workbook->{summary} );
is($result, $target, $caption);


###############################################################################
#
# Test 11. Same as previous + Creation date.
#

# Wed Aug 20 00:20:13 2008
# $sec,$min,$hour,$mday,$mon,$year
# We normalise the time using timegm() so that the tests don't fail due to
# different timezones.
$filetime   = [localtime(timegm(13, 20, 23, 19, 7, 108))];

$workbook->set_properties(
                            title       => 'Title',
                            subject     => 'Subject',
                            author      => 'Author',
                            keywords    => 'Keywords',
                            comments    => 'Comments',
                            last_author => 'Username',
                            created     => $filetime,
                         );

$caption    = " \tset_properties('+ Creation date')";
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

$result     = unpack_record( $workbook->{summary} );
is($result, $target, $caption);


###############################################################################
#
# Test 12. Same as previous. Date set at the workbook level.
#

# Wed Aug 20 00:20:13 2008
# $sec,$min,$hour,$mday,$mon,$year
# We normalise the time using timegm() so that the tests don't fail due to
# different timezones.
$workbook->{_localtime}  = [localtime(timegm(13, 20, 23, 19, 7, 108))];

$workbook->set_properties(
                            title       => 'Title',
                            subject     => 'Subject',
                            author      => 'Author',
                            keywords    => 'Keywords',
                            comments    => 'Comments',
                            last_author => 'Username',
                         );

$caption    = " \tset_properties('+ Creation date')";
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

$result     = unpack_record( $workbook->{summary} );
is($result, $target, $caption);


###############################################################################
#
# Test 13. Same as 11  but params passed as a hashref.
#

# Wed Aug 20 00:20:13 2008
# $sec,$min,$hour,$mday,$mon,$year
# We normalise the time using timegm() so that the tests don't fail due to
# different timezones.
$filetime   = [localtime(timegm(13, 20, 23, 19, 7, 108))];

$workbook->set_properties({
                            title       => 'Title',
                            subject     => 'Subject',
                            author      => 'Author',
                            keywords    => 'Keywords',
                            comments    => 'Comments',
                            last_author => 'Username',
                            created     => $filetime,
                         });

$caption    = " \tset_properties({hash})";
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

$result     = unpack_record( $workbook->{summary} );
is($result, $target, $caption);


###############################################################################
#
# Test 14. UTF-8 string used.
#
SKIP: {

skip " \tset_properties(utf8). Test requires Perl 5.8 Unicode support.", 1
     if $] < 5.008;

$workbook->set_properties(
                            title       => 'Title' . $smiley,
                            created     => undef,
                         );

$caption    = " \tset_properties(utf8)";
$target     = join " ",  qw(
                            FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00
                            00 00 00 00 00 00 00 00 01 00 00 00 E0 85 9F F2
                            F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00
                            34 00 00 00 02 00 00 00 01 00 00 00 18 00 00 00
                            02 00 00 00 20 00 00 00 02 00 00 00 E9 FD 00 00
                            1E 00 00 00 09 00 00 00 54 69 74 6C 65 E2 98 BA
                            00 00 00 00
                           );

$result     = unpack_record( $workbook->{summary} );
is($result, $target, $caption);
}


###############################################################################
#
# Test 15. Manual UTF-8 string used..
#

my $smiley_manual = pack 'H*', 'E298BA';

$workbook->set_properties(
                            title       => 'Title' . $smiley_manual,
                            subject     => 'Subject',
                            created     => undef,
                            utf8        => 1,
                         );

$caption    = " \tset_properties(utf8)";
$target     = join " ",  qw(
                            FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00
                            00 00 00 00 00 00 00 00 01 00 00 00 E0 85 9F F2
                            F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00
                            4C 00 00 00 03 00 00 00 01 00 00 00 20 00 00 00
                            02 00 00 00 28 00 00 00 03 00 00 00 3C 00 00 00
                            02 00 00 00 E9 FD 00 00 1E 00 00 00 09 00 00 00
                            54 69 74 6C 65 E2 98 BA 00 00 00 00 1E 00 00 00
                            08 00 00 00 53 75 62 6A 65 63 74 00
                           );

$result     = unpack_record( $workbook->{summary} );
is($result, $target, $caption);


###############################################################################
#
# Test 16. UTF-8 string used.
#
SKIP: {

skip " \tset_properties(utf8). Test requires Perl 5.8 Unicode support.", 1
     if $] < 5.008;

$workbook->set_properties(
                            title       => 'Title' . $smiley,
                            subject     => 'Subject',
                            created     => undef,
                         );

$caption    = " \tset_properties(utf8)";
$target     = join " ",  qw(
                            FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00
                            00 00 00 00 00 00 00 00 01 00 00 00 E0 85 9F F2
                            F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00
                            4C 00 00 00 03 00 00 00 01 00 00 00 20 00 00 00
                            02 00 00 00 28 00 00 00 03 00 00 00 3C 00 00 00
                            02 00 00 00 E9 FD 00 00 1E 00 00 00 09 00 00 00
                            54 69 74 6C 65 E2 98 BA 00 00 00 00 1E 00 00 00
                            08 00 00 00 53 75 62 6A 65 63 74 00
                           );

$result     = unpack_record( $workbook->{summary} );
is($result, $target, $caption);
}


###############################################################################
#
# Test 17. UTF-8 string used.
#
SKIP: {

skip " \tset_properties(utf8). Test requires Perl 5.8 Unicode support.", 1
     if $] < 5.008;

$workbook->set_properties(
                            title       => 'Title',
                            subject     => 'Subject' . $smiley,
                            created     => undef,
                         );

$caption    = " \tset_properties(utf8)";
$target     = join " ",  qw(
                            FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00
                            00 00 00 00 00 00 00 00 01 00 00 00 E0 85 9F F2
                            F9 4F 68 10 AB 91 08 00 2B 27 B3 D9 30 00 00 00
                            4C 00 00 00 03 00 00 00 01 00 00 00 20 00 00 00
                            02 00 00 00 28 00 00 00 03 00 00 00 38 00 00 00
                            02 00 00 00 E9 FD 00 00 1E 00 00 00 06 00 00 00
                            54 69 74 6C 65 00 00 00 1E 00 00 00 0B 00 00 00
                            53 75 62 6A 65 63 74 E2 98 BA 00 00
                           );

$result     = unpack_record( $workbook->{summary} );
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
