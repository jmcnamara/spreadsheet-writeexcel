#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Tests for the Excel DV structure used in data validation.
#
# reverse('©'), September 2008, John McNamara, jmcnamara@cpan.org
#


use strict;

use Spreadsheet::WriteExcel;
use Test::More tests => 43;


###############################################################################
#
# Tests setup
#
my $test_file           = 'temp_test_file.xls';
my $workbook            = Spreadsheet::WriteExcel->new($test_file);
my $worksheet           = $workbook->add_worksheet();
my $target;
my $result;
my $caption;

# Store worksheet data in memory so we can access it for testing.
$worksheet->{_using_tmpfile} = 0;


###############################################################################
#
# Test 1 Integer between 1 and 10.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: integer between";
$target     = join " ",  qw(
    BE 01 2C 00 01 01 0C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 03
    00 00 00 1E 0A 00 01 00 04 00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 2 Integer not between 1 and 10.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'integer',
        criteria        => 'not between',
        minimum         => 1,
        maximum         => 10,
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: integer not between";
$target     = join " ",  qw(
    BE 01 2C 00 01 01 1C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 03
    00 00 00 1E 0A 00 01 00 04 00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 3,4,5 Integer == 1.
#
for my $operator ('equal to', '=', '==') {
    $worksheet->{_data}         = '';
    $worksheet->{_validations}  = [];

    $worksheet->data_validation('B5',
        {
            validate        => 'integer',
            criteria        => $operator,
            value         => 1,
        });

    $worksheet->_store_validations();

    $caption    = " \tData validation api: integer equal to";
    $target     = join " ",  qw(
        BE 01 29 00 01 01 2C 00 01 00 00 00 01 00 00 00
        01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 00
        00 00 00 01 00 04 00 04 00 01 00 01 00
    );

    $result     = unpack_record($worksheet->{_data});
    is($result, $target, $caption);
}


###############################################################################
#
# Test 6,7,8 Integer != 1.
#
for my $operator ('not equal to', '<>', '!=') {
    $worksheet->{_data}         = '';
    $worksheet->{_validations}  = [];

    $worksheet->data_validation('B5',
        {
            validate        => 'integer',
            criteria        => $operator,
            value         => 1,
        });

    $worksheet->_store_validations();

    $caption    = " \tData validation api: integer not equal to";
    $target     = join " ",  qw(
        BE 01 29 00 01 01 3C 00 01 00 00 00 01 00 00 00
        01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 00
        00 00 00 01 00 04 00 04 00 01 00 01 00
    );

    $result     = unpack_record($worksheet->{_data});
    is($result, $target, $caption);
}


###############################################################################
#
# Test 9,10 Integer > 1.
#
for my $operator ('greater than', '>') {
    $worksheet->{_data}         = '';
    $worksheet->{_validations}  = [];

    $worksheet->data_validation('B5',
        {
            validate        => 'integer',
            criteria        => $operator,
            value         => 1,
        });

    $worksheet->_store_validations();

    $caption    = " \tData validation api: integer >";
    $target     = join " ",  qw(
        BE 01 29 00 01 01 4C 00 01 00 00 00 01 00 00 00
        01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 00
        00 00 00 01 00 04 00 04 00 01 00 01 00
    );

    $result     = unpack_record($worksheet->{_data});
    is($result, $target, $caption);
}


###############################################################################
#
# Test 11,12 Integer < 1.
#
for my $operator ('less than', '<') {
    $worksheet->{_data}         = '';
    $worksheet->{_validations}  = [];

    $worksheet->data_validation('B5',
        {
            validate        => 'integer',
            criteria        => $operator,
            value         => 1,
        });

    $worksheet->_store_validations();

    $caption    = " \tData validation api: integer <";
    $target     = join " ",  qw(
        BE 01 29 00 01 01 5C 00 01 00 00 00 01 00 00 00
        01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 00
        00 00 00 01 00 04 00 04 00 01 00 01 00
    );

    $result     = unpack_record($worksheet->{_data});
    is($result, $target, $caption);
}


###############################################################################
#
# Test 13,14 Integer >= 1.
#
for my $operator ('greater than or equal to', '>=') {
    $worksheet->{_data}         = '';
    $worksheet->{_validations}  = [];

    $worksheet->data_validation('B5',
        {
            validate        => 'integer',
            criteria        => $operator,
            value         => 1,
        });

    $worksheet->_store_validations();

    $caption    = " \tData validation api: integer >=";
    $target     = join " ",  qw(
        BE 01 29 00 01 01 6C 00 01 00 00 00 01 00 00 00
        01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 00
        00 00 00 01 00 04 00 04 00 01 00 01 00
    );

    $result     = unpack_record($worksheet->{_data});
    is($result, $target, $caption);
}


###############################################################################
#
# Test 15,16 Integer <= 1.
#
for my $operator ('less than or equal to', '<=') {
    $worksheet->{_data}         = '';
    $worksheet->{_validations}  = [];

    $worksheet->data_validation('B5',
        {
            validate        => 'integer',
            criteria        => $operator,
            value         => 1,
        });

    $worksheet->_store_validations();

    $caption    = " \tData validation api: integer <=";
    $target     = join " ",  qw(
        BE 01 29 00 01 01 7C 00 01 00 00 00 01 00 00 00
        01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 00
        00 00 00 01 00 04 00 04 00 01 00 01 00
    );

    $result     = unpack_record($worksheet->{_data});
    is($result, $target, $caption);
}


###############################################################################
#
# Test 17 Integer between 1 and 10 (same as test 1) + Ignore blank off.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
        ignore_blank    => 0,
    });

$worksheet->_store_validations();

    $caption    = " \tData validation api: ignore blank off";
$target     = join " ",  qw(
    BE 01 2C 00 01 00 0C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 03
    00 00 00 1E 0A 00 01 00 04 00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 18 Integer between 1 and 10 (same as test 1) + Error style == warning..
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
        error_type      => 'warning',
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: error style = warning";
$target     = join " ",  qw(
        BE 01 2C 00 11 01 0C 00 01 00 00 00 01 00 00 00
        01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 03
        00 00 00 1E 0A 00 01 00 04 00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 19 Integer between 1 and 10 (same as test 1) + Error style == infor..
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
        error_type      => 'information',
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: error style = information";
$target     = join " ",  qw(
        BE 01 2C 00 21 01 0C 00 01 00 00 00 01 00 00 00
        01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 03
        00 00 00 1E 0A 00 01 00 04 00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);

###############################################################################
#
# Test 20 Integer between 1 and 10 (same as test 1)
#         + input title.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
        input_title     => 'Input title January',
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: with input title";
$target     = join " ",  qw(
        BE 01 3E 00 01 01 0C 00 13 00 00 49 6E 70 75 74
        20 74 69 74 6C 65 20 4A 61 6E 75 61 72 79 01 00
        00 00 01 00 00 00 01 00 00 00 03 00 00 00 1E 01
        00 03 00 00 00 1E 0A 00 01 00 04 00 04 00 01 00
        01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 21 Integer between 1 and 10 (same as test 1)
#         + input title.
#         + input message.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
        input_title     => 'Input title January',
        input_message   => 'Input message February',
    });

$worksheet->_store_validations();

$caption    = " \tData validation api:   + input message";
$target     = join " ",  qw(
        BE 01 53 00 01 01 0C 00 13 00 00 49 6E 70 75 74
        20 74 69 74 6C 65 20 4A 61 6E 75 61 72 79 01 00
        00 00 16 00 00 49 6E 70 75 74 20 6D 65 73 73 61
        67 65 20 46 65 62 72 75 61 72 79 01 00 00 00 03
        00 00 00 1E 01 00 03 00 00 00 1E 0A 00 01 00 04
        00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 22 Integer between 1 and 10 (same as test 1)
#         + input title.
#         + input message.
#         + error title.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
        input_title     => 'Input title January',
        input_message   => 'Input message February',
        error_title     => 'Error title March',
    });

$worksheet->_store_validations();

$caption    = " \tData validation api:   + error title";
$target     = join " ",  qw(
        BE 01 63 00 01 01 0C 00 13 00 00 49 6E 70 75 74
        20 74 69 74 6C 65 20 4A 61 6E 75 61 72 79 11 00
        00 45 72 72 6F 72 20 74 69 74 6C 65 20 4D 61 72
        63 68 16 00 00 49 6E 70 75 74 20 6D 65 73 73 61
        67 65 20 46 65 62 72 75 61 72 79 01 00 00 00 03
        00 00 00 1E 01 00 03 00 00 00 1E 0A 00 01 00 04
        00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 23 Integer between 1 and 10 (same as test 1)
#         + input title.
#         + input message.
#         + error title.
#         + error message.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
        input_title     => 'Input title January',
        input_message   => 'Input message February',
        error_title     => 'Error title March',
        error_message   => 'Error message April',
    });

$worksheet->_store_validations();

$caption    = " \tData validation api:   + error message";
$target     = join " ",  qw(
    BE 01 75 00 01 01 0C 00 13 00 00 49 6E 70 75 74
    20 74 69 74 6C 65 20 4A 61 6E 75 61 72 79 11 00
    00 45 72 72 6F 72 20 74 69 74 6C 65 20 4D 61 72
    63 68 16 00 00 49 6E 70 75 74 20 6D 65 73 73 61
    67 65 20 46 65 62 72 75 61 72 79 13 00 00 45 72
    72 6F 72 20 6D 65 73 73 61 67 65 20 41 70 72 69
    6C 03 00 00 00 1E 01 00 03 00 00 00 1E 0A 00 01
    00 04 00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 24 Integer between 1 and 10 (same as test 1)
#         + input title.
#         + input message.
#         + error title.
#         + error message.
#         - input message box.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate            => 'integer',
        criteria            => 'between',
        minimum             => 1,
        maximum             => 10,
        input_title         => 'Input title January',
        input_message       => 'Input message February',
        error_title         => 'Error title March',
        error_message       => 'Error message April',
        show_input          => 0,
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: no input box";
$target     = join " ",  qw(
    BE 01 75 00 01 01 08 00 13 00 00 49 6E 70 75 74
    20 74 69 74 6C 65 20 4A 61 6E 75 61 72 79 11 00
    00 45 72 72 6F 72 20 74 69 74 6C 65 20 4D 61 72
    63 68 16 00 00 49 6E 70 75 74 20 6D 65 73 73 61
    67 65 20 46 65 62 72 75 61 72 79 13 00 00 45 72
    72 6F 72 20 6D 65 73 73 61 67 65 20 41 70 72 69
    6C 03 00 00 00 1E 01 00 03 00 00 00 1E 0A 00 01
    00 04 00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);

###############################################################################
#
# Test 25 Integer between 1 and 10 (same as test 1)
#         + input title.
#         + input message.
#         + error title.
#         + error message.
#         - input message box.
#         - error message box.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate            => 'integer',
        criteria            => 'between',
        minimum             => 1,
        maximum             => 10,
        input_title         => 'Input title January',
        input_message       => 'Input message February',
        error_title         => 'Error title March',
        error_message       => 'Error message April',
        show_input          => 0,
        show_error          => 0,
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: no error box";
$target     = join " ",  qw(
    BE 01 75 00 01 01 00 00 13 00 00 49 6E 70 75 74
    20 74 69 74 6C 65 20 4A 61 6E 75 61 72 79 11 00
    00 45 72 72 6F 72 20 74 69 74 6C 65 20 4D 61 72
    63 68 16 00 00 49 6E 70 75 74 20 6D 65 73 73 61
    67 65 20 46 65 62 72 75 61 72 79 13 00 00 45 72
    72 6F 72 20 6D 65 73 73 61 67 65 20 41 70 72 69
    6C 03 00 00 00 1E 01 00 03 00 00 00 1E 0A 00 01
    00 04 00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 26 'Any' value shouldn't produce a DV record.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'any',
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: any validation";
$target     = '';

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);

###############################################################################
#
# Test 27 Decimal = 1.2345
#
$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'decimal',
        criteria        => '==',
        value         => 1.2345,
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: decimal validation";
$target     = join " ",  qw(
    BE 01 2F 00 02 01 2C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 09 00 00 00 1F 8D 97 6E
    12 83 C0 F3 3F 00 00 00 00 01 00 04 00 04 00 01
    00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 28 List = a,bb,ccc
#
$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'list',
        source          => ['a', 'bb', 'ccc'],
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: explicit list";
$target     = join " ",  qw(
    BE 01 31 00 83 01 0C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 0B 00 00 00 17 08 00 61
    00 62 62 00 63 63 63 00 00 00 00 01 00 04 00 04
    00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 29 List = a,bb,ccc, No dropdown
#
$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'list',
        source          => ['a', 'bb', 'ccc'],
        dropdown        => 0,
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: list with no dropdown";
$target     = join " ",  qw(
    BE 01 31 00 83 03 0C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 0B 00 00 00 17 08 00 61
    00 62 62 00 63 63 63 00 00 00 00 01 00 04 00 04
    00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 30 List = $D$1:$D$5
#
$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('A1:A1',
    {
        validate        => 'list',
        source          => '=$D$1:$D$5',
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: list with range";
$target     = join " ",  qw(
    BE 01 2F 00 03 01 0C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 09 00 00 00 25 00 00 04
    00 03 00 03 00 00 00 00 00 01 00 00 00 00 00 00
    00 00 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 31 Date = 39653 (2008-07-24)
#
$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'date',
        criteria        => '==',
        value           => 39653,
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: date";
$target     = join " ",  qw(
    BE 01 29 00 04 01 2C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 03 00 00 00 1E E5 9A 00
    00 00 00 01 00 04 00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 32 Date = 2008-07-24T
#
$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'date',
        criteria        => '==',
        value           => '2008-07-24T',
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: date auto";
$target     = join " ",  qw(
    BE 01 29 00 04 01 2C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 03 00 00 00 1E E5 9A 00
    00 00 00 01 00 04 00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 33 Date between ranges.
#
$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'date',
        criteria        => 'between',
        minimum         => '2008-01-01T',
        maximum         => '2008-12-12T',
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: date auto, between";
$target     = join " ",  qw(
    BE 01 2C 00 04 01 0C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 03 00 00 00 1E 18 9A 03
    00 00 00 1E 72 9B 01 00 04 00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 34 Time = 0.5 (12:00:00)
#
$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5:B5',
    {
        validate        => 'time',
        criteria        => '==',
        value           => 0.5,
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: time";
$target     = join " ",  qw(
    BE 01 2F 00 05 01 2C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 09 00 00 00 1F 00 00 00
    00 00 00 E0 3F 00 00 00 00 01 00 04 00 04 00 01
    00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 35 Time = T12:00:00
#
$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'time',
        criteria        => '==',
        value           => 'T12:00:00',
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: time auto";
$target     = join " ",  qw(
    BE 01 2F 00 05 01 2C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 09 00 00 00 1F 00 00 00
    00 00 00 E0 3F 00 00 00 00 01 00 04 00 04 00 01
    00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 36 Custom == 10.
#
$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'Custom',
        criteria        => '==',
        value           => 10,
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: custom";
$target     = join " ",  qw(
    BE 01 29 00 07 01 0C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 03 00 00 00 1E 0A 00 00
    00 00 00 01 00 04 00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 37 Check the row/col processing: single A1 style cell.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5',
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: range options";
$target     = join " ",  qw(
    BE 01 2C 00 01 01 0C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 03
    00 00 00 1E 0A 00 01 00 04 00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 38 Check the row/col processing: single A1 style range.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation('B5:B10',
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: range options";
$target     = join " ",  qw(
    BE 01 2C 00 01 01 0C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 03
    00 00 00 1E 0A 00 01 00 04 00 09 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 39 Check the row/col processing: single (row, col) style cell.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation(4, 1,
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: range options";
$target     = join " ",  qw(
    BE 01 2C 00 01 01 0C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 03
    00 00 00 1E 0A 00 01 00 04 00 04 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 40 Check the row/col processing: single (row, col) style range.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation(4, 1, 9, 1,
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: range options";
$target     = join " ",  qw(
    BE 01 2C 00 01 01 0C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 03
    00 00 00 1E 0A 00 01 00 04 00 09 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 41 Check the row/col processing: multiple (row, col) style cells.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation(4, 1,
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
        other_cells     => [[4, 3, 4, 3]],
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: range options";
$target     = join " ",  qw(
    BE 01 34 00 01 01 0C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 03
    00 00 00 1E 0A 00 02 00 04 00 04 00 01 00 01 00
    04 00 04 00 03 00 03 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 42 Check the row/col processing: multiple (row, col) style cells.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation(4, 1,
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
        other_cells     => [[6, 1, 6, 1], [8, 1, 8, 1]],
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: range options";
$target     = join " ",  qw(
    BE 01 3C 00 01 01 0C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 03
    00 00 00 1E 0A 00 03 00 04 00 04 00 01 00 01 00
    06 00 06 00 01 00 01 00 08 00 08 00 01 00 01 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


###############################################################################
#
# Test 43 Check the row/col processing: multiple (row, col) style cells.
#

$worksheet->{_data}         = '';
$worksheet->{_validations}  = [];

$worksheet->data_validation(4, 1, 9, 1,
    {
        validate        => 'integer',
        criteria        => 'between',
        minimum         => 1,
        maximum         => 10,
        other_cells     => [[4, 3, 4, 3]],
    });

$worksheet->_store_validations();

$caption    = " \tData validation api: range options";
$target     = join " ",  qw(
    BE 01 34 00 01 01 0C 00 01 00 00 00 01 00 00 00
    01 00 00 00 01 00 00 00 03 00 00 00 1E 01 00 03
    00 00 00 1E 0A 00 02 00 04 00 09 00 01 00 01 00
    04 00 04 00 03 00 03 00
);

$result     = unpack_record($worksheet->{_data});
is($result, $target, $caption);


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




Tool completed successfully
