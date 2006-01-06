#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Tests to ensure merge formats aren't used in non-merged cells and
# vice-versa. This is temporary feature to prevent users from inadvertently
# making this error.
#
# reverse('©'), April 2005, John McNamara, jmcnamara@cpan.org
#


use strict;

use Spreadsheet::WriteExcel;
use Test::More tests => 8;


###############################################################################
#
# Tests setup
#
my $test_file           = "temp_test_file.xls";
my $workbook            = Spreadsheet::WriteExcel->new($test_file);
my $worksheet           = $workbook->add_worksheet();
my $merged_format       = $workbook->add_format(bold => 1);
my $non_merged_format   = $workbook->add_format(bold => 1);


$worksheet->set_row(5, undef, $merged_format);
$worksheet->set_column('G:G', undef, $merged_format);

###############################################################################
#
# Test
#
eval {
    $worksheet->write      ('A1',    'Test', $non_merged_format);
    $worksheet->merge_range('A3:B4', 'Test', $merged_format    );
};
ok(! $@, " \tNormal usage.");


###############################################################################
#
# Test
#
eval {
    $worksheet->write      ('D1',    'Test', $merged_format    );
};
ok(  $@, " \tMerge format in non-merged cell.");


###############################################################################
#
# Test
#
eval {
    $worksheet->merge_range('D3:E4', 'Test', $non_merged_format);
};
ok(  $@, " \tNon merge format in merged cells.");


###############################################################################
#
# Test
#
eval {
    $worksheet->write('G1', 'Test',);
};
ok(  $@, " \tMerge format in column.");


###############################################################################
#
# Test
#
eval {
    $worksheet->write('A6', 'Test',);
};
ok(  $@, " \tMerge format in row.");


###############################################################################
#
# Test
#
eval {
    $worksheet->write('G6', 'Test',);
};
ok(  $@, " \tMerge format in column and row.");


###############################################################################
#
# Test
#
eval {
    $worksheet->write('H7', 'Test',);
};
ok(! $@, " \tNo merge format in column and row.");

###############################################################################
#
# Test
#
eval {
    $worksheet->write      ('A1',    'Test', $non_merged_format);
    $worksheet->merge_range('A3:B4', 'Test', $merged_format    );
};
ok(! $@, " \tNormal usage again.");




$workbook->close();
unlink $test_file;


__END__



