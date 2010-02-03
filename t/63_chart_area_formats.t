###############################################################################
#
# A test for Spreadsheet::WriteExcel::Chart.
#
# Tests for the set*area() Chart methods.
#
# reverse('©'), January 2010, John McNamara, jmcnamara@cpan.org
#

# prove -I../lib --nocolor -v 63_chart_area_formats.t

use strict;

use Spreadsheet::WriteExcel;

use Test::More tests => 52;
#use Test::More 'no_plan';


###############################################################################
#
# Tests setup
#
my $test_file = 'temp_test_file.xls';
my $workbook  = Spreadsheet::WriteExcel->new( $test_file );
my $chart     = $workbook->add_chart( type => 'column' );
$chart->{_using_tmpfile} = 0;

my $got_line;
my $got_area;
my $expected_line;
my $expected_area;
my $caption1 = " \tChart: chartarea format - line";
my $caption2 = " \tChart: chartarea format - area";


###############################################################################
#
# 1. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_chartarea(
    color        => 'red',
    line_color   => 'black',
    line_pattern => 2,
    line_weight  => 3,
);

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 01 00 01 00 00 00 08 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 3. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_chartarea( color => 'red', );

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 05 00 FF FF 08 00 4D 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 5. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_chartarea( line_color => 'red', );

$expected_line = join ' ', qw(
  07 10 0C 00 FF 00 00 00 00 00 FF FF 00 00 0A 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF FF FF 00 00 00 00 00 00 00 00 00
  4E 00 4D 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 7. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_chartarea( line_pattern => 2, );

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 01 00 FF FF 00 00 4F 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF FF FF 00 00 00 00 00 00 00 00 00
  4E 00 4D 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 9. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_chartarea( line_weight => 3, );

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 00 00 01 00 00 00 4F 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF FF FF 00 00 00 00 00 00 00 00 00
  4E 00 4D 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 11. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_chartarea(
    color      => 'red',
    line_color => 'black',
);

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 00 00 FF FF 00 00 08 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 13. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_chartarea(
    color        => 'red',
    line_pattern => 2,
);

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 01 00 FF FF 00 00 4F 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 15. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_chartarea(
    color       => 'red',
    line_weight => 3,
);

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 00 00 01 00 00 00 4F 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# Set some variables for embedded chart tests.
#
$caption1           = " \tChart: embedded chartarea format - line";
$caption2           = " \tChart: embedded chartarea format - area";
$chart->{_embedded} = 1;


###############################################################################
#
# 17. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart, 1 );

$chart->set_chartarea();

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 00 00 00 00 09 00 4D 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF FF FF 00 00 00 00 00 01 00 01 00
  4E 00 4D 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 19. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart, 1 );

$chart->set_chartarea(
    color        => 'red',
    line_color   => 'black',
    line_pattern => 2,
    line_weight  => 3,
);

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 01 00 01 00 00 00 08 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 21. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart, 1 );

$chart->set_chartarea( color => 'red', );

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 00 00 FF FF 09 00 4D 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 23. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart, 1 );

$chart->set_chartarea( line_color => 'red', );

$expected_line = join ' ', qw(
  07 10 0C 00 FF 00 00 00 00 00 FF FF 00 00 0A 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF FF FF 00 00 00 00 00 01 00 01 00
  4E 00 4D 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 25. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart, 1 );

$chart->set_chartarea( line_pattern => 2, );

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 01 00 FF FF 00 00 4F 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF FF FF 00 00 00 00 00 01 00 01 00
  4E 00 4D 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 27. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart, 1 );

$chart->set_chartarea( line_weight => 3, );

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 00 00 01 00 00 00 4F 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF FF FF 00 00 00 00 00 01 00 01 00
  4E 00 4D 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 29. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart, 1 );

$chart->set_chartarea(
    color      => 'red',
    line_color => 'black',
);

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 00 00 FF FF 00 00 08 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 31. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart, 1 );

$chart->set_chartarea(
    color        => 'red',
    line_pattern => 2,
);

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 01 00 FF FF 00 00 4F 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 33. Test the chartarea format methods. See the set_*area() properties below.
#
reset_chart( $chart, 1 );

$chart->set_chartarea(
    color       => 'red',
    line_weight => 3,
);

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 00 00 01 00 00 00 4F 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_chartarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# Set some variables for embedded chart tests.
#
$caption1           = " \tChart: plotarea format - line";
$caption2           = " \tChart: plotarea format - area";
$chart->{_embedded} = 0;


###############################################################################
#
# 35. Test the plotarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_plotarea();

$expected_line = join ' ', qw(
  07 10 0C 00 80 80 80 00 00 00 00 00 00 00 17 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 C0 C0 C0 00 00 00 00 00 01 00 00 00
  16 00 4F 00
);

( $got_line, $got_area ) = get_plotarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 37. Test the plotarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_plotarea(
    color        => 'red',
    line_color   => 'black',
    line_pattern => 2,
    line_weight  => 3,
);

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 01 00 01 00 00 00 08 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_plotarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 39. Test the plotarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_plotarea( color => 'red', );

$expected_line = join ' ', qw(
  07 10 0C 00 80 80 80 00 00 00 00 00 00 00 17 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_plotarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 41. Test the plotarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_plotarea( line_color => 'red', );

$expected_line = join ' ', qw(
  07 10 0C 00 FF 00 00 00 00 00 00 00 00 00 0A 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 C0 C0 C0 00 00 00 00 00 01 00 00 00
  16 00 08 00
);

( $got_line, $got_area ) = get_plotarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 43. Test the plotarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_plotarea( line_pattern => 2, );

$expected_line = join ' ', qw(
  07 10 0C 00 80 80 80 00 01 00 00 00 00 00 17 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 C0 C0 C0 00 00 00 00 00 01 00 00 00
  16 00 08 00
);

( $got_line, $got_area ) = get_plotarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 45. Test the plotarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_plotarea( line_weight => 3, );

$expected_line = join ' ', qw(
  07 10 0C 00 80 80 80 00 00 00 01 00 00 00 17 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 C0 C0 C0 00 00 00 00 00 01 00 00 00
  16 00 08 00
);

( $got_line, $got_area ) = get_plotarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 47. Test the plotarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_plotarea(
    color      => 'red',
    line_color => 'black',
);

$expected_line = join ' ', qw(
  07 10 0C 00 00 00 00 00 00 00 00 00 00 00 08 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_plotarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 49. Test the plotarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_plotarea(
    color        => 'red',
    line_pattern => 2,
);

$expected_line = join ' ', qw(
  07 10 0C 00 80 80 80 00 01 00 00 00 00 00 17 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_plotarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# 51. Test the plotarea format methods. See the set_*area() properties below.
#
reset_chart( $chart );

$chart->set_plotarea(
    color       => 'red',
    line_weight => 3,
);

$expected_line = join ' ', qw(
  07 10 0C 00 80 80 80 00 00 00 01 00 00 00 17 00
);

$expected_area = join ' ', qw(
  0A 10 10 00 FF 00 00 00 00 00 00 00 01 00 00 00
  0A 00 08 00
);

( $got_line, $got_area ) = get_plotarea_formats( $chart );

is( $got_line, $expected_line, $caption1 );
is( $got_area, $expected_area, $caption2 );


###############################################################################
#
# Utility functions used by the test suite.
#
###############################################################################


###############################################################################
#
# Reset the chart data for testing.
#
sub reset_chart {

    my $chart    = shift;
    my $embedded = shift;

    # Reset the chart data.
    $chart->{_data} = '';
    $chart->_set_default_properties();

    if ( $embedded ) {
        $chart->_set_embedded_config_data();
    }
}


###############################################################################
#
# Extract Line and Area format records from the Chartarea Frame stream.
#
sub get_chartarea_formats {

    $chart = shift;

    $chart->_store_chartarea_frame_stream();

    my $line = unpack_record( substr $chart->{_data}, 12, 16 );
    my $area = unpack_record( substr $chart->{_data}, 28, 20 );

    return ( $line, $area );
}


###############################################################################
#
# Extract Line and Area format records from the Chartarea Frame stream.
#
sub get_plotarea_formats {

    $chart = shift;

    $chart->_store_plotarea_frame_stream();

    my $line = unpack_record( substr $chart->{_data}, 12, 16 );
    my $area = unpack_record( substr $chart->{_data}, 28, 20 );

    return ( $line, $area );
}


###############################################################################
#
# Unpack the binary data into a format suitable for printing in tests.
#
sub unpack_record {
    return join ' ', map { sprintf '%02X', $_ } unpack 'C*', shift;
}


###############################################################################
#
# Clean up.
#
$workbook->close();
unlink $test_file;


__END__
