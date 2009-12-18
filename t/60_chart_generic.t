###############################################################################
#
# A test for Spreadsheet::WriteExcel::Chart.
#
# Tests for the Excel Chart.pm methods.
#
# reverse('©'), December 2009, John McNamara, jmcnamara@cpan.org
#

# prove -I../lib --nocolor -v 60_chart_generic.t

use strict;

use Spreadsheet::WriteExcel::Chart;

#use Test::More tests => 12;
use Test::More 'no_plan';


###############################################################################
#
# Tests setup
#
my $chart = Spreadsheet::WriteExcel::Chart->new();
my $got;
my $expected;
my $caption;
my @values;

###############################################################################
#
# Test the _store_fbi method.
#
$caption = " \tChart: _store_fbi()";

$expected = join ' ', qw(
  60 10 0A 00 B8 38 A1 22 C8 00 00 00 05 00
);

$got = unpack_record( $chart->_store_fbi(5) );

is( $got, $expected, $caption );

###############################################################################
#
# Test the _store_fbi method.
#
$expected = join ' ', qw(
  60 10 0A 00 B8 38 A1 22 C8 00 00 00 06 00
);

$got = unpack_record( $chart->_store_fbi(6) );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_chart method.
#
$caption = " \tChart: _store_chart()";

$expected = join ' ', qw(
    02 10 10 00 00 00 00 00 00 00 00 00 E0 51 DD 02
    38 B8 C2 01
);

$got = unpack_record( $chart->_store_chart() );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_series() method.
#
$caption = " \tChart: _store_series()";

$expected = join ' ', qw(
    03 10 0C 00 01 00 01 00 08 00 08 00 01 00 00 00
);

@values = ( 1, 1, 8, 8, 1, 0 );

$got = unpack_record( $chart->_store_series(@values) );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_begin() method.
#
$caption = " \tChart: _store_begin()";

$expected = join ' ', qw(
     33 10 00 00
);

$got = unpack_record( $chart->_store_begin() );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_end() method.
#
$caption = " \tChart: _store_end()";

$expected = join ' ', qw(
     34 10 00 00
);

$got = unpack_record( $chart->_store_end() );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_ai() method.
#
$caption = " \tChart: _store_ai()";

@values = ( 0, 1, 0, '' );

$expected = join ' ', qw(
    51 10 08 00 00 01 00 00 00 00 00 00
);

$got = unpack_record( $chart->_store_ai( @values ) );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_ai() method.
#
$caption = " \tChart: _store_ai()";

@values = ( 1, 2, 0, pack 'H*', '3B00000000070000000000' );

$expected = join ' ', qw(
    51 10 13 00 01 02 00 00 00 00 0B 00 3B 00 00 00
    00 07 00 00 00 00 00
);

$got = unpack_record( $chart->_store_ai( @values ) );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_dataformat() method.
#
$caption = " \tChart: _store_dataformat()";

$expected = join ' ', qw(
    06 10 08 00 FF FF 00 00 00 00 00 00
);

$got = unpack_record( $chart->_store_dataformat( 0 ) );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_3dbarshape() method.
#
$caption = " \tChart: _store_3dbarshape()";

$expected = join ' ', qw(
    5F 10 02 00 00 00
);

$got = unpack_record( $chart->_store_3dbarshape() );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_sertocrt() method.
#
$caption = " \tChart: _store_sertocrt()";

$expected = join ' ', qw(
    45 10 02 00 00 00
);

$got = unpack_record( $chart->_store_sertocrt() );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_shtprops() method.
#
$caption = " \tChart: _store_shtprops()";

$expected = join ' ', qw(
    44 10 04 00 0E 00 00 00
);

$got = unpack_record( $chart->_store_shtprops() );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_defaulttext() method.
#
$caption = " \tChart: _store_defaulttext()";

$expected = join ' ', qw(
    24 10 02 00 02 00
);

$got = unpack_record( $chart->_store_defaulttext() );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_chart_text() method.
#
$caption = " \tChart: _store_chart_text()";

$expected = join ' ', qw(
    25 10 20 00 02 02 01 00 00 00 00 00 EA FF FF FF
    DC FF FF FF 00 00 00 00 00 00 00 00 B1 00 4D 00
    20 10 00 00
);

$got = unpack_record( $chart->_store_chart_text() );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_fontx() method.
#
$caption = " \tChart: _store_fontx()";

$expected = join ' ', qw(
    26 10 02 00 05 00
);

$got = unpack_record( $chart->_store_fontx( 5 ) );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_axesused() method.
#
$caption = " \tChart: _store_axesused()";

$expected = join ' ', qw(
    46 10 02 00 01 00
);

$got = unpack_record( $chart->_store_axesused( 1 ) );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_axisparent() method.
#
$caption = " \tChart: _store_axisparent()";

$expected = join ' ', qw(
    41 10 12 00 00 00 A0 00 00 00 99 00 00 00 B2 0D
    00 00 E4 0D 00 00
);

$got = unpack_record( $chart->_store_axisparent( 0 ) );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_axis() method.
#
$caption = " \tChart: _store_axis()";

$expected = join ' ', qw(
    1D 10 12 00 00 00 00 00 00 00 00 00 00 00 00 00
    00 00 00 00 00 00
);

$got = unpack_record( $chart->_store_axis( 0 ) );

is( $got, $expected, $caption );








###############################################################################
#
# Utility function used by the test suite.
#
# Unpack the binary data into a format suitable for printing in tests.
#
sub unpack_record {
    return join ' ', map { sprintf '%02X', $_ } unpack 'C*', shift;
}



__END__
