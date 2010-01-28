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

use Test::More tests => 20;
#use Test::More 'no_plan';


###############################################################################
#
# Tests setup
#
my $test_file = 'temp_test_file.xls';
my $workbook  = Spreadsheet::WriteExcel->new( $test_file );
my $chart     = $workbook->add_chart( type => 'column' );
my %values;
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
# Test. Line patterns with indices.
#
$caption = " \tChart: \$pattern = _get_line_pattern()";

%values = (
    0     => 5,
    1     => 0,
    2     => 1,
    3     => 2,
    4     => 3,
    5     => 4,
    6     => 7,
    7     => 6,
    8     => 8,
    9     => 0,
    undef => 0
);

$expected = [];
$got      = [];

while ( my ( $user, $excel ) = each %values ) {
    push @$got,      $chart->_get_line_pattern( $user );
    push @$expected, $excel;
}

is_deeply( $got, $expected, $caption );


###############################################################################
#
# Test. Line patterns with names.
#
$caption = " \tChart: \$pattern = _get_line_pattern()";

%values = (
    'solid'        => 0,
    'dash'         => 1,
    'dot'          => 2,
    'dash-dot'     => 3,
    'dash-dot-dot' => 4,
    'none'         => 5,
    'dark-gray'    => 6,
    'medium-gray'  => 7,
    'light-gray'   => 8,
    'DASH'         => 1,
    'fictional'    => 0
);

$expected = [];
$got      = [];

while ( my ( $user, $excel ) = each %values ) {
    push @$got,      $chart->_get_line_pattern( $user );
    push @$expected, $excel;
}

is_deeply( $got, $expected, $caption );


###############################################################################
#
# Test. Line weights with indices.
#
$caption = " \tChart: \$weight  = _get_line_weight()";

%values = (
    1     => -1,
    2     => 0,
    3     => 1,
    4     => 2,
    5     => 0,
    0     => 0,
    undef => 0
);

$expected = [];
$got      = [];

while ( my ( $user, $excel ) = each %values ) {
    push @$got,      $chart->_get_line_weight( $user );
    push @$expected, $excel;
}

is_deeply( $got, $expected, $caption );


###############################################################################
#
# Test. Line weights with names.
#
$caption = " \tChart: \$weight  = _get_line_weight()";

%values = (
    'hairline'  => -1,
    'narrow'    => 0,
    'medium'    => 1,
    'wide'      => 2,
    'WIDE'      => 2,
    'Fictional' => 0,
);

$expected = [];
$got      = [];

while ( my ( $user, $excel ) = each %values ) {
    push @$got,      $chart->_get_line_weight( $user );
    push @$expected, $excel;
}

is_deeply( $got, $expected, $caption );


###############################################################################
#
# Clean up.
#
$workbook->close();
unlink $test_file;


__END__
