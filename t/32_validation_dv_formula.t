#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Tests for the packed formula strings used in the Excel DV structure
# as part of data validation.
#
# reverse('©'), September 2008, John McNamara, jmcnamara@cpan.org
#


use strict;

use Spreadsheet::WriteExcel;
use Test::More tests => 12;


###############################################################################
#
# Tests setup
#
my $test_file           = "temp_test_file.xls";
my $workbook            = Spreadsheet::WriteExcel->new($test_file);
my $worksheet           = $workbook->add_worksheet();
my $worksheet2          = $workbook->add_worksheet();
my $target;
my $result;
my $caption;

my $formula;
my @bytes;



###############################################################################
#
# Test 1 Integer values.
#
$formula      = '10';

$caption    = " \tData validation: _pack_dv_formula('$formula')";
@bytes      = qw(
                    03 00 00 E0 1E 0A 00
                );

# Zero out Excel's random unused word to allow comparison.
$bytes[2]   = '00';
$bytes[3]   = '00';
$target     = join " ", @bytes;

$result     = unpack_record($worksheet->_pack_dv_formula($formula));
is($result, $target, $caption);


###############################################################################
#
# Test 2 Decimal values.
#
$formula      = '1.2345';

$caption    = " \tData validation: _pack_dv_formula('$formula')";
@bytes      = qw(
                    09 00 E0 3F 1F 8D 97 6E 12 83 C0 F3 3F
                );

# Zero out Excel's random unused word to allow comparison.
$bytes[2]   = '00';
$bytes[3]   = '00';
$target     = join " ", @bytes;

$result     = unpack_record($worksheet->_pack_dv_formula($formula));
is($result, $target, $caption);


###############################################################################
#
# Test 3 Date values..
#
$formula      = $worksheet->convert_date_time('2008-07-24T');

$caption    = " \tData validation: _pack_dv_formula('2008-07-24')";
@bytes      = qw(
                    03 00 E0 3F 1E E5 9A
                );

# Zero out Excel's random unused word to allow comparison.
$bytes[2]   = '00';
$bytes[3]   = '00';
$target     = join " ", @bytes;

$result     = unpack_record($worksheet->_pack_dv_formula($formula));
is($result, $target, $caption);


###############################################################################
#
# Test 4 Time values.
#
$formula      = $worksheet->convert_date_time('T12:00');

$caption    = " \tData validation: _pack_dv_formula('12:00')";
@bytes      = qw(
                    09 00 E0 3F 1F 00 00 00 00 00 00 E0 3F
                );

# Zero out Excel's random unused word to allow comparison.
$bytes[2]   = '00';
$bytes[3]   = '00';
$target     = join " ", @bytes;

$result     = unpack_record($worksheet->_pack_dv_formula($formula));
is($result, $target, $caption);


###############################################################################
#
# Test 5 Cell reference value.
#
$formula      = '=C9';

$caption    = " \tData validation: _pack_dv_formula('$formula')";
@bytes      = qw(
                    05 00 E0 3F 44 08 00 02 C0
                );

# Zero out Excel's random unused word to allow comparison.
$bytes[2]   = '00';
$bytes[3]   = '00';
$target     = join " ", @bytes;

$result     = unpack_record($worksheet->_pack_dv_formula($formula));
is($result, $target, $caption);


###############################################################################
#
# Test 6 Cell reference value.
#
$formula      = '=E3:E6';

$caption    = " \tData validation: _pack_dv_formula('$formula')";
@bytes      = qw(
                    09 00 0C 00 25 02 00 05 00 04 C0 04 C0
                );

# Zero out Excel's random unused word to allow comparison.
$bytes[2]   = '00';
$bytes[3]   = '00';
$target     = join " ", @bytes;

$result     = unpack_record($worksheet->_pack_dv_formula($formula));
is($result, $target, $caption);


###############################################################################
#
# Test 7 Cell reference value.
#
$formula      = '=$E$3:$E$6';

$caption    = " \tData validation: _pack_dv_formula('$formula')";
@bytes      = qw(
                    09 00 0C 00 25 02 00 05 00 04 00 04 00
                );

# Zero out Excel's random unused word to allow comparison.
$bytes[2]   = '00';
$bytes[3]   = '00';
$target     = join " ", @bytes;

$result     = unpack_record($worksheet->_pack_dv_formula($formula));
is($result, $target, $caption);


###############################################################################
#
# Test 8 Cell reference value.
#
$formula      = '=$E$3:$E$6';

$caption    = " \tData validation: _pack_dv_formula('$formula')";
@bytes      = qw(
                    09 00 0C 00 25 02 00 05 00 04 00 04 00
                );

# Zero out Excel's random unused word to allow comparison.
$bytes[2]   = '00';
$bytes[3]   = '00';
$target     = join " ", @bytes;

$result     = unpack_record($worksheet->_pack_dv_formula($formula));
is($result, $target, $caption);


###############################################################################
#
# Test 9 List values.
#
$formula      = ['a', 'bb', 'ccc'];

$caption    = " \tData validation: _pack_dv_formula(['a', 'bb', 'ccc'])";
@bytes      = qw(
                    0B 00 0C 00 17 08 00 61 00 62 62 00 63 63 63
                );

# Zero out Excel's random unused word to allow comparison.
$bytes[2]   = '00';
$bytes[3]   = '00';
$target     = join " ", @bytes;

$result     = unpack_record($worksheet->_pack_dv_formula($formula));
is($result, $target, $caption);


###############################################################################
#
# Test 10 Empty string.
#
$formula      = '';

$caption    = " \tData validation: _pack_dv_formula('')";
@bytes      = qw(
                    00 00 00
                );

# Zero out Excel's random unused word to allow comparison.
$bytes[2]   = '00';
$bytes[3]   = '00';
$target     = join " ", @bytes;

$result     = unpack_record($worksheet->_pack_dv_formula($formula));
is($result, $target, $caption);


###############################################################################
#
# Test 11 Undefined value.
#
$formula      = undef;

$caption    = " \tData validation: _pack_dv_formula(undef)";
@bytes      = qw(
                    00 00 00
                );

# Zero out Excel's random unused word to allow comparison.
$bytes[2]   = '00';
$bytes[3]   = '00';
$target     = join " ", @bytes;

$result     = unpack_record($worksheet->_pack_dv_formula($formula));
is($result, $target, $caption);




###############################################################################
#
# Test 10 List values (with a utf8 string).
#
SKIP: {

skip " \_pack_dv_string(). Test requires Perl 5.8 Unicode support.", 1
     if $] < 5.008;

my $euro    = chr 0x20Ac;
$formula    = ['a', 'bb', 'ccc', $euro];

$caption    = " \tData validation: _pack_dv_formula(['a', 'bb', 'ccc', utf8])";
@bytes      = qw(
                    17 00 0C 00 17 0A 01 61 00 00 00 62 00 62 00 00 00
                    63 00 63 00 63 00 00 00 AC 20
                );

# Zero out Excel's random unused word to allow comparison.
$bytes[2]   = '00';
$bytes[3]   = '00';
$target     = join " ", @bytes;

$result     = unpack_record($worksheet->_pack_dv_formula($formula));
is($result, $target, $caption);

}

# TODO
# Test failing reference to Sheet2!A1
# Test for formula string > 255 chars
# $formula      = ['a' x 256];




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



