#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Tests for the token extraction method used to parse autofilter expressions.
#
# reverse('©'), August 2007, John McNamara, jmcnamara@cpan.org
#


use strict;

use Spreadsheet::WriteExcel;
use Test::More tests => 18;


###############################################################################
#
# Tests setup
#
my $test_file = "temp_test_file.xls";
my $workbook   = Spreadsheet::WriteExcel->new($test_file);
my $worksheet  = $workbook->add_worksheet();


###############################################################################
#
# Test cases structured as [$input, [@expected_output]]
#
my @tests = (
    [
        undef,
        [],
    ],

    [
        '',
        [],
    ],

    [
        '0 <  2000',
        [0, '<', 2000],
    ],

    [
        'x <  2000',
        ['x', '<', 2000],
    ],

    [
        'x >  2000',
        ['x', '>', 2000],
    ],

    [
        'x == 2000',
        ['x', '==', 2000],
    ],

    [
        'x >  2000 and x <  5000',
        ['x', '>',  2000, 'and', 'x', '<', 5000],
    ],

    [
        'x = "foo"',
        ['x', '=', 'foo'],
    ],

    [
        'x = foo',
        ['x', '=', 'foo'],
    ],

    [
        'x = "foo bar"',
        ['x', '=', 'foo bar'],
    ],

    [
        'x = "foo "" bar"',
        ['x', '=', 'foo " bar'],
    ],

    [
        'x = "foo bar" or x = "bar foo"',
        ['x', '=', 'foo bar', 'or', 'x', '=', 'bar foo'],
    ],

    [
        'x = "foo "" bar" or x = "bar "" foo"',
        ['x', '=', 'foo " bar', 'or', 'x', '=', 'bar " foo'],
    ],

    [
        'x = """"""""',
        ['x', '=', '"""'],
    ],

    [
        'x = Blanks',
        ['x', '=', 'Blanks'],
    ],

    [
        'x = NonBlanks',
        ['x', '=', 'NonBlanks'],
    ],

    [
        'top 10 %',
        ['top', 10, '%'],
    ],

    [
        'top 10 items',
        ['top', 10, 'items'],
    ],

);


###############################################################################
#
# Run the test cases.
#
for my $aref (@tests) {
    my $expression  = $aref->[0];
    my $expected    = $aref->[1];
    my @results     = $worksheet->_extract_filter_tokens($expression);

    my $testname    = $expression || 'none';

    is_deeply(\@results, $expected, " \t" . $testname);
}

# Cleanup
$workbook->close();
unlink $test_file;

__END__
