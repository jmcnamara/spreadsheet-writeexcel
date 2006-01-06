#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Tests for some of the internal method used to write the NOTE record that
# is used in cell comments.
#
# reverse('©'), September 2005, John McNamara, jmcnamara@cpan.org
#


use strict;

use Spreadsheet::WriteExcel;
use Test::More tests => 5;


###############################################################################
#
# Tests setup
#
my $test_file           = "temp_test_file.xls";
my $workbook            = Spreadsheet::WriteExcel->new($test_file);
my $worksheet           = $workbook->add_worksheet();
my $target;
my $result;
my $caption;
my $row;
my $col;
my $obj_id;
my $visible;
my $author;
my $encoding;
my @data;



###############################################################################
#
# Test 1 NOTE. Blank author name.
#

@data = $worksheet->_comment_params(2, 0, 'Test');

$row        = $data[0];
$col        = $data[1];
$author     = $data[4];
$encoding   = $data[5];
$visible    = $data[6];
$obj_id     = 1;

$caption    = " \t_store_note()";
$target     = join " ",  qw(
                            1C 00 0C 00 02 00 00 00 00 00 01 00 00 00 00 00
                           );


$result     = unpack_record($worksheet->_store_note($row,
                                                    $col,
                                                    $obj_id,
                                                    $author,
                                                    $encoding,
                                                    $visible,
                                                    ));
is($result, $target, $caption);


###############################################################################
#
# Test 2 NOTE. Defined author name
#

@data = $worksheet->_comment_params(2, 0, 'Test', author => 'Username');

$row        = $data[0];
$col        = $data[1];
$author     = $data[4];
$encoding   = $data[5];
$visible    = $data[6];
$obj_id     = 1;

$caption    = " \t_store_note()";
$target     = join " ",  qw(
                            1C 00 14 00 02 00 00 00 00 00 01 00 08 00 00 55
                            73 65 72 6E 61 6D 65 00
                           );


$result     = unpack_record($worksheet->_store_note($row,
                                                    $col,
                                                    $obj_id,
                                                    $author,
                                                    $encoding,
                                                    $visible,
                                                    ));
is($result, $target, $caption);


###############################################################################
#
# Test 3 NOTE. Visible note.
#

@data = $worksheet->_comment_params(4, 2, 'Test', author  => 'Username',
                                                  visible => 1
                                    );

$row        = $data[0];
$col        = $data[1];
$author     = $data[4];
$encoding   = $data[5];
$visible    = $data[6];
$obj_id     = 1;


$caption    = " \t_store_note()";
$target     = join " ",  qw(
                            1C 00 14 00 04 00 02 00 02 00 01 00 08 00 00 55
                            73 65 72 6E 61 6D 65 00
                           );


$result     = unpack_record($worksheet->_store_note($row,
                                                    $col,
                                                    $obj_id,
                                                    $author,
                                                    $encoding,
                                                    $visible,
                                                    ));
is($result, $target, $caption);


###############################################################################
#
# Test 3 NOTE. UTF16 author name.
#

$author     = pack "n", 0x20Ac; # Euro symbol

@data = $worksheet->_comment_params(4, 2, 'Test', author          =>$author,
                                                  author_encoding => 1
                                    );

$row        = $data[0];
$col        = $data[1];
$author     = $data[4];
$encoding   = $data[5];
$visible    = $data[6];
$obj_id     = 1;

$caption    = " \t_store_note()";
$target     = join " ",  qw(
                            1C 00 0E 00 04 00 02 00 00 00 01 00 01 00 01 AC
                            20 00
                           );


$result     = unpack_record($worksheet->_store_note($row,
                                                    $col,
                                                    $obj_id,
                                                    $author,
                                                    $encoding,
                                                    $visible,
                                                    ));
is($result, $target, $caption);


###############################################################################
#
# Test 4 NOTE. UTF8 author name. Perl 5.8 only.
#

SKIP: {

skip " \t_store_note() skipped test requires Perl 5.8 Unicode support", 1
     if $] < 5.008;

$author     = chr 0x20Ac; # Euro symbol

@data       = $worksheet->_comment_params(4, 2, 'Test',  author =>$author);

$row        = $data[0];
$col        = $data[1];
$author     = $data[4];
$encoding   = $data[5];
$visible    = $data[6];
$obj_id     = 1;


$caption    = " \t_store_note()";
$target     = join " ",  qw(
                            1C 00 0E 00 04 00 02 00 00 00 01 00 01 00 01 AC
                            20 00
                           );


$result     = unpack_record($worksheet->_store_note($row,
                                                    $col,
                                                    $obj_id,
                                                    $author,
                                                    $encoding,
                                                    $visible,
                                                    ));
is($result, $target, $caption);

}


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



