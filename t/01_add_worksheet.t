#!/usr/bin/perl -wl

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Tests for valid worksheet name handling.
#
# reverse('©'), March 2005, John McNamara, jmcnamara@cpan.org
#


use strict;

use Spreadsheet::WriteExcel;
use Test::More tests => 69;


# Tests for valid and invalid worksheet names
my @tests1  = (
                # Tests for valid names
                [ 'PASS', undef,      'No worksheet name'           ],
                [ 'PASS', '',         'Blank worksheet name'        ],
                [ 'PASS', 'Sheet10',  'Valid worksheet name'        ],
                [ 'PASS', 'a' x 31,   'Valid 31 char name'          ],

                # Tests for invalid names
                [ 'FAIL', 'Sheet1',   'Caught duplicate name'       ],
                [ 'FAIL', 'Sheet2',   'Caught duplicate name'       ],
                [ 'FAIL', 'Sheet3',   'Caught duplicate name'       ],
                [ 'FAIL', 'sheet1',   'Caught case-insensitive name'],
                [ 'FAIL', 'SHEET1',   'Caught case-insensitive name'],
                [ 'FAIL', 'sheetz',   'Caught case-insensitive name'],
                [ 'FAIL', 'SHEETZ',   'Caught case-insensitive name'],
                [ 'FAIL', 'a' x 32,   'Caught long name'            ],
                [ 'FAIL', '[',        'Caught invalid char'         ],
                [ 'FAIL', ']',        'Caught invalid char'         ],
                [ 'FAIL', ':',        'Caught invalid char'         ],
                [ 'FAIL', '*',        'Caught invalid char'         ],
                [ 'FAIL', '?',        'Caught invalid char'         ],
                [ 'FAIL', '/',        'Caught invalid char'         ],
                [ 'FAIL', '\\',       'Caught invalid char'         ],

             );




###############################################################################
#
# Tests 1. ASCII tests
#
my $test_file  = "temp_test_file.xml";
my $workbook   = Spreadsheet::WriteExcel->new($test_file);
my $worksheet1 = $workbook->add_worksheet();        # Implicit name 'Sheet1'
my $worksheet2 = $workbook->add_worksheet();        # Implicit name 'Sheet2'
my $worksheet3 = $workbook->add_worksheet('Sheet3');
my $worksheet4 = $workbook->add_worksheet('Sheetz');

for my $test_ref (@tests1) {

    my $target    = $test_ref->[0];
    my $sheetname = $test_ref->[1];
    my $caption   = $test_ref->[2];

    eval {$workbook->_check_sheetname($sheetname)};

    my $result = $@ ? 'FAIL' : 'PASS';

    $sheetname = 'undef' unless defined $sheetname;

    is($result, $target, sprintf " \t%-7s %-28s: %s",
                                 'ASCII:', $caption, $sheetname);
}

$workbook->close();


###############################################################################
#
# Tests 2. UTF16-BE tests
#

$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();        # Implicit name 'Sheet1'
$worksheet2 = $workbook->add_worksheet();        # Implicit name 'Sheet2'
$worksheet3 = $workbook->add_worksheet('Sheet3');
$worksheet4 = $workbook->add_worksheet("\0S\0h\0e\0e\0t\0z", 1);

for my $test_ref (@tests1) {

    my $target    = $test_ref->[0];
    my $sheetname = $test_ref->[1];
    my $caption   = $test_ref->[2];

    # Convert ASCII to UTF16-BE if not blank or undef
    $sheetname = pack "n*", unpack "C*", $sheetname if $sheetname;

    eval {$workbook->_check_sheetname($sheetname, 1)};

    my $result = $@ ? 'FAIL' : 'PASS';

    $sheetname = 'undef' unless defined $sheetname;

    # Change null byte to \0 for printing
    $sheetname =~ s/\0/\\0/g;

    is($result, $target, sprintf " \t%-7s %-28s: %s",
                                 'UTF-16:', $caption, $sheetname);
}

$workbook->close();



###############################################################################
#
# Tests 3. UTF-8 tests
#

SKIP: {


my $uni = chr 0x263A;
my @tests2  = (
                # Tests for valid names
                [ 'PASS', $uni,      'Unicode char'                 ],
                [ 'PASS', $uni x 31,   'Valid 31 char name'         ],

                # Tests for invalid names
                [ 'FAIL', chr 0x0438, 'Caught duplicate name'       ],
                [ 'FAIL', chr 0x0418, 'Caught case-insensitive name'],
                [ 'FAIL', $uni x 32,  'Caught long name'            ],
                [ 'FAIL', '[' . $uni, 'Caught invalid char'         ],
                [ 'FAIL', ']' . $uni, 'Caught invalid char'         ],
                [ 'FAIL', ':' . $uni, 'Caught invalid char'         ],
                [ 'FAIL', '*' . $uni, 'Caught invalid char'         ],
                [ 'FAIL', '?' . $uni, 'Caught invalid char'         ],
                [ 'FAIL', '/' . $uni, 'Caught invalid char'         ],
                [ 'FAIL', '\\'. $uni, 'Caught invalid char'         ],

             );

skip "\tskipped tests requires Perl 5.8 Unicode support", 0 + @tests1 + @tests2 if $] < 5.008;


$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();        # Implicit name 'Sheet1'
$worksheet2 = $workbook->add_worksheet();        # Implicit name 'Sheet2'
$worksheet3 = $workbook->add_worksheet('Sheet3');
$worksheet4 = $workbook->add_worksheet("\0S\0h\0e\0e\0t\0z", 1);
my $worksheet5 = $workbook->add_worksheet(chr 0x0438);


for my $test_ref (@tests1) {

    my $target    = $test_ref->[0];
    my $sheetname = $test_ref->[1];
    my $caption   = $test_ref->[2];

    require Encode;
    $sheetname = Encode::encode_utf8($sheetname) if $sheetname;

    eval {$workbook->_check_sheetname($sheetname)};

    my $result = $@ ? 'FAIL' : 'PASS';

    $sheetname = 'undef' unless defined $sheetname;

    # Change null byte to \0 for printing
    $sheetname =~ s/\0/\\0/g;

    is($result, $target, sprintf " \t%-7s %-28s: %s",
                                 'UTF-8:', $caption, $sheetname);
}


for my $test_ref (@tests2) {

    my $target    = $test_ref->[0];
    my $sheetname = $test_ref->[1];
    my $caption   = $test_ref->[2];

    eval {$workbook->_check_sheetname($sheetname)};

    my $result = $@ ? 'FAIL' : 'PASS';

    $sheetname = 'undef' unless defined $sheetname;

    is($result, $target, sprintf " \t%-7s %-28s: %s",
                                 'UTF-8:', $caption, '');
}


$workbook->close();


}

unlink $test_file;


__END__



