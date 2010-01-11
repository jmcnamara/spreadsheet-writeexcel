###############################################################################
#
# A test for Spreadsheet::WriteExcel::Chart subclass methods.
#
# Tests for the Excel Chart.pm methods.
#
# reverse('©'), December 2009, John McNamara, jmcnamara@cpan.org
#

# prove -I../lib --nocolor -v 61_chart_subclasses.t

use strict;

use Spreadsheet::WriteExcel::Chart;

use Test::More tests => 7;


###############################################################################
#
# Tests setup
#
my $chart;
my $got;
my $expected;
my $caption;


###############################################################################
#
# Test for overridden _store_chart_type() in ::Chart::Column.pm.
#
$chart = Spreadsheet::WriteExcel::Chart->factory( 'column' );

$caption = " \tChart: Column()";

$expected = join ' ', qw(
  17 10 06 00 00 00 96 00 00 00
);

$got = unpack_record( $chart->_store_chart_type() );

is( $got, $expected, $caption );


###############################################################################
#
# Test for overridden _store_chart_type() in ::Chart::Bar.pm.
#
$chart = Spreadsheet::WriteExcel::Chart->factory( 'bar' );

$caption = " \tChart: Bar()";

$expected = join ' ', qw(
  17 10 06 00 00 00 96 00 01 00
);

$got = unpack_record( $chart->_store_chart_type() );

is( $got, $expected, $caption );


###############################################################################
#
# Test for overridden _store_chart_type() in ::Chart::Line.pm.
#
$chart = Spreadsheet::WriteExcel::Chart->factory( 'line' );

$caption = " \tChart: Line()";

$expected = join ' ', qw(
  18 10 02 00 00 00
);

$got = unpack_record( $chart->_store_chart_type() );

is( $got, $expected, $caption );


###############################################################################
#
# Test for overridden _store_chart_type() in ::Chart::Area.pm.
#
$chart = Spreadsheet::WriteExcel::Chart->factory( 'area' );

$caption = " \tChart: Area()";

$expected = join ' ', qw(
  1A 10 02 00 01 00
);

$got = unpack_record( $chart->_store_chart_type() );

is( $got, $expected, $caption );


###############################################################################
#
# Test for overridden _store_chart_type() in ::Chart::Pie.pm.
#
$chart = Spreadsheet::WriteExcel::Chart->factory( 'pie' );

$caption = " \tChart: Pie()";

$expected = join ' ', qw(
  19 10 06 00 00 00 00 00 02 00
);

$got = unpack_record( $chart->_store_chart_type() );

is( $got, $expected, $caption );


###############################################################################
#
# Test for overridden _store_chart_type() in ::Chart::Scatter.pm.
#
$chart = Spreadsheet::WriteExcel::Chart->factory( 'scatter' );

$caption = " \tChart: Scatter()";

$expected = join ' ', qw(
  1B 10 06 00 64 00 01 00 00 00
);

$got = unpack_record( $chart->_store_chart_type() );

is( $got, $expected, $caption );


###############################################################################
#
# Test for overridden _store_chart_type() in ::Chart::Stock.pm.
#
$chart = Spreadsheet::WriteExcel::Chart->factory( 'stock' );

$caption = " \tChart: Stock()";

$expected = join ' ', qw(
  18 10 02 00 00 00
);

$got = unpack_record( $chart->_store_chart_type() );

is( $got, $expected, $caption );


##############################################################################
#
# Utility function used by the test suite.
#
# Unpack the binary data into a format suitable for printing in tests.
#
sub unpack_record {
    return join ' ', map { sprintf '%02X', $_ } unpack 'C*', shift;
}


__END__
