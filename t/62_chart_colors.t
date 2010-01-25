###############################################################################
#
# A test for Spreadsheet::WriteExcel::Chart.
#
# Tests for the Excel Chart.pm colour conversion methods.
#
# reverse('©'), January 2010, John McNamara, jmcnamara@cpan.org
#

# prove -I../lib --nocolor -v 62_chart_colors.t

use strict;

use Spreadsheet::WriteExcel;

use Test::More tests => 2;


###############################################################################
#
# Tests setup
#
my $test_file = 'temp_test_file.xls';
my $workbook  = Spreadsheet::WriteExcel->new( $test_file );
my $chart     = $workbook->add_chart( type => 'column' );
my $color;
my $got_index;
my $got_rgb;
my $expected_index;
my $expected_rgb;
my $caption1;
my $caption2;


###############################################################################
#
# Test 1.  User defined colour as string.
#
$color    = 'red';
$caption1 = " \tChart: \$index = _get_color_indices( $color )";
$caption2 = " \tChart: \$rgb   = _get_color_indices( $color )";

$expected_index = 0x0A;
$expected_rgb   = 0x000000FF;


( $got_index, $got_rgb ) = $chart->_get_color_indices( $color );

is( $got_index, $expected_index, $caption1 );
is( $got_rgb,   $expected_rgb,   $caption2 );


# Clean up.
$workbook->close();
unlink $test_file;


__END__
