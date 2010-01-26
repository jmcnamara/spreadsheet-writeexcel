###############################################################################
#
# A test for Spreadsheet::WriteExcel::Chart.
#
# Tests for the Excel Chart.pm format conversion methods.
#
# reverse('©'), January 2010, John McNamara, jmcnamara@cpan.org
#

# prove -I../lib --nocolor -v 62_chart_formats.t

use strict;

use Spreadsheet::WriteExcel;

use Test::More tests => 18;
#use Test::More 'no_plan';


###############################################################################
#
# Tests setup
#
my $test_file = 'temp_test_file.xls';
my $workbook  = Spreadsheet::WriteExcel->new( $test_file );
my $chart     = $workbook->add_chart( type => 'column' );
my @values;
my $color;
my $got;
my $got_index;
my $got_rgb;
my $expected;
my $expected_index;
my $expected_rgb;
my $caption;
my $caption1;
my $caption2;


###############################################################################
#
# Test. User defined colour as string.
#
$color    = 'red';
$caption1 = " \tChart: \$index   = _get_color_indices( $color )";
$caption2 = " \tChart: \$rgb     = _get_color_indices( $color )";

$expected_index = 0x0A;
$expected_rgb   = 0x000000FF;


( $got_index, $got_rgb ) = $chart->_get_color_indices( $color );

is( $got_index, $expected_index, $caption1 );
is( $got_rgb,   $expected_rgb,   $caption2 );


###############################################################################
#
# Test. User defined colour as string.
#
$color    = 'black';
$caption1 = " \tChart: \$index   = _get_color_indices( $color )";
$caption2 = " \tChart: \$rgb     = _get_color_indices( $color )";

$expected_index = 0x08;
$expected_rgb   = 0x00000000;


( $got_index, $got_rgb ) = $chart->_get_color_indices( $color );

is( $got_index, $expected_index, $caption1 );
is( $got_rgb,   $expected_rgb,   $caption2 );


###############################################################################
#
# Test. User defined colour as string.
#
$color    = 'white';
$caption1 = " \tChart: \$index   = _get_color_indices( $color )";
$caption2 = " \tChart: \$rgb     = _get_color_indices( $color )";

$expected_index = 0x09;
$expected_rgb   = 0x00FFFFFF;


( $got_index, $got_rgb ) = $chart->_get_color_indices( $color );

is( $got_index, $expected_index, $caption1 );
is( $got_rgb,   $expected_rgb,   $caption2 );


###############################################################################
#
# Test. User defined colour as an index.
#
$color    = 0x0A;
$caption1 = " \tChart: \$index   = _get_color_indices( $color )";
$caption2 = " \tChart: \$rgb     = _get_color_indices( $color )";

$expected_index = 0x0A;
$expected_rgb   = 0x000000FF;


( $got_index, $got_rgb ) = $chart->_get_color_indices( $color );

is( $got_index, $expected_index, $caption1 );
is( $got_rgb,   $expected_rgb,   $caption2 );


###############################################################################
#
# Test. User defined colour as an out of range index.
#
$color    = 7;
$caption1 = " \tChart: \$index   = _get_color_indices( $color )";
$caption2 = " \tChart: \$rgb     = _get_color_indices( $color )";

$expected_index = undef;
$expected_rgb   = undef;


( $got_index, $got_rgb ) = $chart->_get_color_indices( $color );

is( $got_index, $expected_index, $caption1 );
is( $got_rgb,   $expected_rgb,   $caption2 );


###############################################################################
#
# Test. User defined colour as an out of range index.
#
$color    = 64;
$caption1 = " \tChart: \$index   = _get_color_indices( $color )";
$caption2 = " \tChart: \$rgb     = _get_color_indices( $color )";

$expected_index = undef;
$expected_rgb   = undef;


( $got_index, $got_rgb ) = $chart->_get_color_indices( $color );

is( $got_index, $expected_index, $caption1 );
is( $got_rgb,   $expected_rgb,   $caption2 );


###############################################################################
#
# Test. User defined colour as an invalid string.
#
$color    = 'plaid';
$caption1 = " \tChart: \$index   = _get_color_indices( $color )";
$caption2 = " \tChart: \$rgb     = _get_color_indices( $color )";

$expected_index = undef;
$expected_rgb   = undef;


( $got_index, $got_rgb ) = $chart->_get_color_indices( $color );

is( $got_index, $expected_index, $caption1 );
is( $got_rgb,   $expected_rgb,   $caption2 );


###############################################################################
#
# Test. User defined colour as an undef property.
#
$color    = undef;
$caption1 = " \tChart: \$index   = _get_color_indices( undef )";
$caption2 = " \tChart: \$rgb     = _get_color_indices( undef )";

$expected_index = undef;
$expected_rgb   = undef;


( $got_index, $got_rgb ) = $chart->_get_color_indices( $color );

is( $got_index, $expected_index, $caption1 );
is( $got_rgb,   $expected_rgb,   $caption2 );


###############################################################################
#
# Test. Line patterns
#
$caption = " \tChart: \$pattern = _get_line_pattern()";

@values   = ( 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, undef );
$expected = [ 5, 0, 1, 2, 3, 4, 7, 6, 8, 1, 1     ];
$got      = [];

for my $pattern ( @values ) {
    push @$got, $chart->_get_line_pattern( $pattern );
}

is_deeply( $got, $expected, $caption );


###############################################################################
#
# Test. Line weights
#
$caption = " \tChart: \$weight  = _get_line_weight()";

@values   = ( 0,  1, 2, 3, 4, 5, undef );
$expected = [ 0, -1, 0, 1, 2, 0, 0     ];
$got      = [];

for my $weight ( @values ) {
    push @$got, $chart->_get_line_weight( $weight );
}

is_deeply( $got, $expected, $caption );



###############################################################################
#
# Clean up.
#
$workbook->close();
unlink $test_file;


__END__
