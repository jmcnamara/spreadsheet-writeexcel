#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Tests for the internal methods used to write the MSODRAWINGGROUP record.
#
# reverse('©'), September 2005, John McNamara, jmcnamara@cpan.org
#


use strict;

use Spreadsheet::WriteExcel;
use Test::More tests => 34;


###############################################################################
#
# Tests setup
#
my $test_file = "temp_test_file.xls";
my $workbook;
my $worksheet1;
my $worksheet2;
my $worksheet3;
my $target;
my $result;
my $caption;
my $count1;
my $count2;
my $count3;
my @target_ids;
my @result_ids;


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$count1     = 1;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 5A 00 0F 00 00 F0 52 00 00 00 00 00 06 F0
                            18 00 00 00 02 04 00 00 02 00 00 00 02 00 00 00
                            01 00 00 00 01 00 00 00 02 00 00 00 33 00 0B F0
                            12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
                            C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
                            00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
              );

$caption    = sprintf " \tSheet1: %4d comments.", $count1;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 2, 1025,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$count1     = 2;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 5A 00 0F 00 00 F0 52 00 00 00 00 00 06 F0
                            18 00 00 00 03 04 00 00 02 00 00 00 03 00 00 00
                            01 00 00 00 01 00 00 00 03 00 00 00 33 00 0B F0
                            12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
                            C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
                            00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments.", $count1;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 3, 1026,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$count1     = 3;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 5A 00 0F 00 00 F0 52 00 00 00 00 00 06 F0
                            18 00 00 00 04 04 00 00 02 00 00 00 04 00 00 00
                            01 00 00 00 01 00 00 00 04 00 00 00 33 00 0B F0
                            12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
                            C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
                            00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments.", $count1;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 4, 1027,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$count1     = 1023;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 5A 00 0F 00 00 F0 52 00 00 00 00 00 06 F0
                            18 00 00 00 00 08 00 00 02 00 00 00 00 04 00 00
                            01 00 00 00 01 00 00 00 00 04 00 00 33 00 0B F0
                            12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
                            C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
                            00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments.", $count1;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 1024, 2047,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$count1     = 1024;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 62 00 0F 00 00 F0 5A 00 00 00 00 00 06 F0
                            20 00 00 00 01 08 00 00 03 00 00 00 01 04 00 00
                            01 00 00 00 01 00 00 00 00 04 00 00 01 00 00 00
                            01 00 00 00 33 00 0B F0 12 00 00 00 BF 00 08 00
                            08 00 81 01 09 00 00 08 C0 01 40 00 00 08 40 00
                            1E F1 10 00 00 00 0D 00 00 08 0C 00 00 08 17 00
                            00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments.", $count1;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 1025, 2048,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$count1     = 2048;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 6A 00 0F 00 00 F0 62 00 00 00 00 00 06 F0
                            28 00 00 00 01 0C 00 00 04 00 00 00 01 08 00 00
                            01 00 00 00 01 00 00 00 00 04 00 00 01 00 00 00
                            00 04 00 00 01 00 00 00 01 00 00 00 33 00 0B F0
                            12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
                            C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
                            00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments.", $count1;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 2049, 3072,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$worksheet2 = $workbook->add_worksheet();
$count1     = 1;
$count2     = 1;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;
$worksheet2->write_comment($_ -1, 0, 'aaa') for 1 .. $count2;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 62 00 0F 00 00 F0 5A 00 00 00 00 00 06 F0
                            20 00 00 00 02 08 00 00 03 00 00 00 04 00 00 00
                            02 00 00 00 01 00 00 00 02 00 00 00 02 00 00 00
                            02 00 00 00 33 00 0B F0 12 00 00 00 BF 00 08 00
                            08 00 81 01 09 00 00 08 C0 01 40 00 00 08 40 00
                            1E F1 10 00 00 00 0D 00 00 08 0C 00 00 08 17 00
                            00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments, Sheet2: %4d comments.",
                      $count1, $count2;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 2, 1025,
                2048, 2, 2, 2049,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$worksheet2 = $workbook->add_worksheet();
$count1     = 2;
$count2     = 2;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;
$worksheet2->write_comment($_ -1, 0, 'aaa') for 1 .. $count2;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 62 00 0F 00 00 F0 5A 00 00 00 00 00 06 F0
                            20 00 00 00 03 08 00 00 03 00 00 00 06 00 00 00
                            02 00 00 00 01 00 00 00 03 00 00 00 02 00 00 00
                            03 00 00 00 33 00 0B F0 12 00 00 00 BF 00 08 00
                            08 00 81 01 09 00 00 08 C0 01 40 00 00 08 40 00
                            1E F1 10 00 00 00 0D 00 00 08 0C 00 00 08 17 00
                            00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments, Sheet2: %4d comments.",
                      $count1, $count2;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 3, 1026,
                2048, 2, 3, 2050,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$worksheet2 = $workbook->add_worksheet();
$count1     = 1023;
$count2     = 1;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;
$worksheet2->write_comment($_ -1, 0, 'aaa') for 1 .. $count2;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 62 00 0F 00 00 F0 5A 00 00 00 00 00 06 F0
                            20 00 00 00 02 08 00 00 03 00 00 00 02 04 00 00
                            02 00 00 00 01 00 00 00 00 04 00 00 02 00 00 00
                            02 00 00 00 33 00 0B F0 12 00 00 00 BF 00 08 00
                            08 00 81 01 09 00 00 08 C0 01 40 00 00 08 40 00
                            1E F1 10 00 00 00 0D 00 00 08 0C 00 00 08 17 00
                            00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments, Sheet2: %4d comments.",
                      $count1, $count2;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 1024, 2047,
                2048, 2, 2, 2049,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$worksheet2 = $workbook->add_worksheet();
$count1     = 1023;
$count2     = 1023;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;
$worksheet2->write_comment($_ -1, 0, 'aaa') for 1 .. $count2;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 62 00 0F 00 00 F0 5A 00 00 00 00 00 06 F0
                            20 00 00 00 00 0C 00 00 03 00 00 00 00 08 00 00
                            02 00 00 00 01 00 00 00 00 04 00 00 02 00 00 00
                            00 04 00 00 33 00 0B F0 12 00 00 00 BF 00 08 00
                            08 00 81 01 09 00 00 08 C0 01 40 00 00 08 40 00
                            1E F1 10 00 00 00 0D 00 00 08 0C 00 00 08 17 00
                            00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments, Sheet2: %4d comments.",
                      $count1, $count2;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 1024, 2047,
                2048, 2, 1024, 3071,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$worksheet2 = $workbook->add_worksheet();
$count1     = 1024;
$count2     = 1024;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;
$worksheet2->write_comment($_ -1, 0, 'aaa') for 1 .. $count2;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 72 00 0F 00 00 F0 6A 00 00 00 00 00 06 F0
                            30 00 00 00 01 10 00 00 05 00 00 00 02 08 00 00
                            02 00 00 00 01 00 00 00 00 04 00 00 01 00 00 00
                            01 00 00 00 02 00 00 00 00 04 00 00 02 00 00 00
                            01 00 00 00 33 00 0B F0 12 00 00 00 BF 00 08 00
                            08 00 81 01 09 00 00 08 C0 01 40 00 00 08 40 00
                            1E F1 10 00 00 00 0D 00 00 08 0C 00 00 08 17 00
                            00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments, Sheet2: %4d comments.",
                      $count1, $count2;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 1025, 2048,
                3072, 2, 1025, 4096,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$worksheet2 = $workbook->add_worksheet();
$count1     = 1024;
$count2     = 1;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;
$worksheet2->write_comment($_ -1, 0, 'aaa') for 1 .. $count2;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 6A 00 0F 00 00 F0 62 00 00 00 00 00 06 F0
                            28 00 00 00 02 0C 00 00 04 00 00 00 03 04 00 00
                            02 00 00 00 01 00 00 00 00 04 00 00 01 00 00 00
                            01 00 00 00 02 00 00 00 02 00 00 00 33 00 0B F0
                            12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
                            C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
                            00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments, Sheet2: %4d comments.",
                      $count1, $count2;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 1025, 2048,
                3072, 2, 2, 3073,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$worksheet2 = $workbook->add_worksheet();
$worksheet3 = $workbook->add_worksheet();
$count1     = 1023;
$count2     = 1;
$count3     = 1023;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;
$worksheet2->write_comment($_ -1, 0, 'aaa') for 1 .. $count2;
$worksheet3->write_comment($_ -1, 0, 'aaa') for 1 .. $count3;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 6A 00 0F 00 00 F0 62 00 00 00 00 00 06 F0
                            28 00 00 00 00 10 00 00 04 00 00 00 02 08 00 00
                            03 00 00 00 01 00 00 00 00 04 00 00 02 00 00 00
                            02 00 00 00 03 00 00 00 00 04 00 00 33 00 0B F0
                            12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
                            C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
                            00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments, Sheet2: %4d comments, " .
                         "Sheet3: %4d comments.", $count1, $count2, $count3;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 1024, 2047,
                2048, 2, 2, 2049,
                3072, 3, 1024, 4095,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$worksheet2 = $workbook->add_worksheet();
$worksheet3 = $workbook->add_worksheet();
$count1     = 1023;
$count2     = 1023;
$count3     = 1;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;
$worksheet2->write_comment($_ -1, 0, 'aaa') for 1 .. $count2;
$worksheet3->write_comment($_ -1, 0, 'aaa') for 1 .. $count3;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 6A 00 0F 00 00 F0 62 00 00 00 00 00 06 F0
                            28 00 00 00 02 0C 00 00 04 00 00 00 02 08 00 00
                            03 00 00 00 01 00 00 00 00 04 00 00 02 00 00 00
                            00 04 00 00 03 00 00 00 02 00 00 00 33 00 0B F0
                            12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
                            C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
                            00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments, Sheet2: %4d comments, " .
                         "Sheet3: %4d comments.", $count1, $count2, $count3;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 1024, 2047,
                2048, 2, 1024, 3071,
                3072, 3, 2, 3073,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$worksheet2 = $workbook->add_worksheet();
$worksheet3 = $workbook->add_worksheet();
$count1     = 1024;
$count2     = 1;
$count3     = 1024;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;
$worksheet2->write_comment($_ -1, 0, 'aaa') for 1 .. $count2;
$worksheet3->write_comment($_ -1, 0, 'aaa') for 1 .. $count3;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 7A 00 0F 00 00 F0 72 00 00 00 00 00 06 F0
                            38 00 00 00 01 14 00 00 06 00 00 00 04 08 00 00
                            03 00 00 00 01 00 00 00 00 04 00 00 01 00 00 00
                            01 00 00 00 02 00 00 00 02 00 00 00 03 00 00 00
                            00 04 00 00 03 00 00 00 01 00 00 00 33 00 0B F0
                            12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
                            C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
                            00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments, Sheet2: %4d comments, " .
                         "Sheet3: %4d comments.", $count1, $count2, $count3;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 1025, 2048,
                3072, 2, 2, 3073,
                4096, 3, 1025, 5120,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$worksheet2 = $workbook->add_worksheet();
$worksheet3 = $workbook->add_worksheet();
$count1     = 1024;
$count2     = 1024;
$count3     = 1;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;
$worksheet2->write_comment($_ -1, 0, 'aaa') for 1 .. $count2;
$worksheet3->write_comment($_ -1, 0, 'aaa') for 1 .. $count3;

$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 7A 00 0F 00 00 F0 72 00 00 00 00 00 06 F0
                            38 00 00 00 02 14 00 00 06 00 00 00 04 08 00 00
                            03 00 00 00 01 00 00 00 00 04 00 00 01 00 00 00
                            01 00 00 00 02 00 00 00 00 04 00 00 02 00 00 00
                            01 00 00 00 03 00 00 00 02 00 00 00 33 00 0B F0
                            12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
                            C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
                            00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments, Sheet2: %4d comments, " .
                         "Sheet3: %4d comments.", $count1, $count2, $count3;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 1025, 2048,
                3072, 2, 1025, 4096,
                5120, 3, 2, 5121,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Test. Same as previous except also tests that duplicates are ignored.
#
$workbook   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1 = $workbook->add_worksheet();
$worksheet2 = $workbook->add_worksheet();
$worksheet3 = $workbook->add_worksheet();
$count1     = 1024;
$count2     = 1024;
$count3     = 1;

$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;
$worksheet2->write_comment($_ -1, 0, 'aaa') for 1 .. $count2;
$worksheet3->write_comment($_ -1, 0, 'aaa') for 1 .. $count3;

# Duplicates.
$worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;
$worksheet2->write_comment($_ -1, 0, 'aaa') for 1 .. $count2;
$worksheet3->write_comment($_ -1, 0, 'aaa') for 1 .. $count3;


$workbook->_calc_mso_sizes();

$target     = join " ",  qw(
                            EB 00 7A 00 0F 00 00 F0 72 00 00 00 00 00 06 F0
                            38 00 00 00 02 14 00 00 06 00 00 00 04 08 00 00
                            03 00 00 00 01 00 00 00 00 04 00 00 01 00 00 00
                            01 00 00 00 02 00 00 00 00 04 00 00 02 00 00 00
                            01 00 00 00 03 00 00 00 02 00 00 00 33 00 0B F0
                            12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
                            C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
                            00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
                           );

$caption    = sprintf " \tSheet1: %4d comments, Sheet2: %4d comments, " .
                         "Sheet3: %4d comments.", $count1, $count2, $count3;

$result     = unpack_record($workbook->_add_mso_drawing_group());
is($result, $target, $caption);


# Test the parameters pass to the worksheets
$caption   .= ' (params)';
@result_ids = ();
@target_ids = (
                1024, 1, 1025, 2048,
                3072, 2, 1025, 4096,
                5120, 3, 2, 5121,
              );

for my $sheet ($workbook->sheets()) {
    push @result_ids, @{$sheet->{_object_ids}};
}

is_deeply(\@result_ids, \@target_ids , $caption);

$workbook->close();


###############################################################################
#
# Unpack the binary data into a format suitable for printing in tests.
#
sub unpack_record {
    return join ' ', map {sprintf "%02X", $_} unpack "C*", $_[0];
}


# Cleanup
unlink $test_file;


__END__



