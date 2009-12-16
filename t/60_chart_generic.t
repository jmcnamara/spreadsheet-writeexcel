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

###############################################################################
#
# Test the _store_fbi method.
#
$caption = " \tChart, _store_fbi()";

$expected = join ' ', qw(
  60 10 0A 00 B8 38 A1 22 C8 00 00 00 05 00
);

$got = unpack_record( $chart->_store_fbi(5) );

is( $got, $expected, $caption );

# Try a different index.
$expected = join ' ', qw(
  60 10 0A 00 B8 38 A1 22 C8 00 00 00 06 00
);

$got = unpack_record( $chart->_store_fbi(6) );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _store_chart method.
#
$caption = " \tChart, _store_chart()";

$expected = join ' ', qw(
    02 10 10 00 00 00 00 00 00 00 00 00 E0 51 DD 02
    38 B8 C2 01
);

$got = unpack_record( $chart->_store_chart() );

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
