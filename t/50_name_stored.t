###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Tests for the Excel EXTERNSHEET and NAME records created by print_are()..
#
# reverse('©'), September 2008, John McNamara, jmcnamara@cpan.org
#


use strict;

use Spreadsheet::WriteExcel;
use Test::More tests => 8;
#use Test::More 'no_plan';


###############################################################################
#
# Tests setup
#
my $test_file = "temp_test_file.xls";
my $workbook   = Spreadsheet::WriteExcel->new($test_file);
my $worksheet  = $workbook->add_worksheet();

my $target;
my $result;
my $caption;
my $name;
my $encoding;
my $sheet_index;
my $formula;


###############################################################################
#
# Test 1. Test for print_area() NAME with simple range.
#

$caption        = " \tNAME for \$worksheet1->print_area('A1:B12')";
$name           = pack 'C', 0x06;
$encoding       = 0;
$sheet_index    = 1;
$formula        = pack 'H*', '3B000000000B0000000100';

$result         = $workbook->_store_name(
                                            $name,
                                            $encoding,
                                            $sheet_index,
                                            $formula
                                        );

$target         = pack 'H*', join '', qw(
    18 00 1B 00 20 00 00 01 0B 00 00 00 01 00 00 00
    00 00 00 06 3B 00 00 00 00 0B 00 00 00 01 00
);

$result         = _unpack_name($result);
$target         = _unpack_name($target);
is_deeply($result, $target, $caption);


###############################################################################
#
# Test 2. Test for print_area() NAME with simple range in sheet 3.
#

$caption        = " \tNAME for \$worksheet3->print_area('G7:H8')";
$name           = pack 'C', 0x06;
$encoding       = 0;
$sheet_index    = 3;
$formula        = pack 'H*', '3B02000600070006000700';

$result         = $workbook->_store_name(
                                            $name,
                                            $encoding,
                                            $sheet_index,
                                            $formula
                                        );

$target         = pack 'H*', join '', qw(
    18 00 1B 00 20 00 00 01 0B 00 00 00 03 00 00 00
    00 00 00 06 3B 02 00 06 00 07 00 06 00 07 00
);

$result         = _unpack_name($result);
$target         = _unpack_name($target);
is_deeply($result, $target, $caption);


###############################################################################
#
# Test 3. Test for repeat_rows() NAME.
#

$caption        = " \tNAME for \$worksheet1->repeat_rows(0, 9)";
$name           = pack 'C', 0x07;
$encoding       = 0;
$sheet_index    = 1;
$formula        = pack 'H*', '3B0000000009000000FF00';

$result         = $workbook->_store_name(
                                            $name,
                                            $encoding,
                                            $sheet_index,
                                            $formula
                                        );

$target         = pack 'H*', join '', qw(
    18 00 1B 00 20 00 00 01 0B 00 00 00 01 00 00 00
    00 00 00 07 3B 00 00 00 00 09 00 00 00 FF 00
);

$result         = _unpack_name($result);
$target         = _unpack_name($target);
is_deeply($result, $target, $caption);

$workbook->close();


###############################################################################
#
# Test 4. Test for repeat_rows() NAME on sheet 3.
#

$caption        = " \tNAME for \$worksheet3->repeat_rows(6, 7)";
$name           = pack 'C', 0x07;
$encoding       = 0;
$sheet_index    = 3;
$formula        = pack 'H*', '3B0200060007000000FF00';

$result         = $workbook->_store_name(
                                            $name,
                                            $encoding,
                                            $sheet_index,
                                            $formula
                                        );

$target         = pack 'H*', join '', qw(
    18 00 1B 00 20 00 00 01 0B 00 00 00 03 00 00 00
    00 00 00 07 3B 02 00 06 00 07 00 00 00 FF 00
);

$result         = _unpack_name($result);
$target         = _unpack_name($target);
is_deeply($result, $target, $caption);


###############################################################################
#
# Test 5. Test for repeat_columns() NAME.
#

$caption        = " \tNAME for \$worksheet1->repeat_columns('A:J')";
$name           = pack 'C', 0x07;
$encoding       = 0;
$sheet_index    = 1;
$formula        = pack 'H*', '3B00000000FFFF00000900';

$result         = $workbook->_store_name(
                                            $name,
                                            $encoding,
                                            $sheet_index,
                                            $formula
                                        );

$target         = pack 'H*', join '', qw(
    18 00 1B 00 20 00 00 01 0B 00 00 00 01 00 00 00
    00 00 00 07 3B 00 00 00 00 FF FF 00 00 09 00
);

$result         = _unpack_name($result);
$target         = _unpack_name($target);
is_deeply($result, $target, $caption);


###############################################################################
#
# Test 6. Test for repeat_rows() and repeat_columns() together NAME.
#

$caption        = " \tNAME for repeat_rows(1, 2) repeat_columns(3, 4)";
$name           = pack 'C', 0x07;
$encoding       = 0;
$sheet_index    = 1;
$formula        = pack 'H*', '2917003B00000000FFFF030004003B0000010002000000FF0010';

$result         = $workbook->_store_name(
                                            $name,
                                            $encoding,
                                            $sheet_index,
                                            $formula
                                        );

$target         = pack 'H*', join '', qw(
    18 00 2A 00 20 00 00 01 1A 00 00 00 01 00 00 00
    00 00 00 07 29 17 00 3B 00 00 00 00 FF FF 03 00
    04 00 3B 00 00 01 00 02 00 00 00 FF 00 10
);

$result         = _unpack_name($result);
$target         = _unpack_name($target);
is_deeply($result, $target, $caption);


###############################################################################
#
# Test 7. Test for print_area() NAME with simple range.
#

$caption        = " \tNAME for \$worksheet1->autofilter('A1:C5');";
$name           = pack 'C', 0x0D;
$encoding       = 0;
$sheet_index    = 1;
$formula        = pack 'H*', '3B00000000040000000200';

$result         = $workbook->_store_name(
                                            $name,
                                            $encoding,
                                            $sheet_index,
                                            $formula
                                        );

$target         = pack 'H*', join '', qw(
    18 00 1B 00 21 00 00 01 0B 00 00 00 01 00 00 00
    00 00 00 0D 3B 00 00 00 00 04 00 00 00 02 00
);

$result         = _unpack_name($result);
$target         = _unpack_name($target);
is_deeply($result, $target, $caption);


###############################################################################
#
# Test 8. Test for define_name() global NAME.
#

$caption        = " \tNAME for \$worksheet1->define_name('Foo', ...);";
$name           = 'Foo';
$encoding       = 0;
$sheet_index    = 0;
$formula        = pack 'H*', '3A000007000100';

$result         = $workbook->_store_name(
                                            $name,
                                            $encoding,
                                            $sheet_index,
                                            $formula
                                        );

$target         = pack 'H*', join '', qw(
    18 00 19 00 00 00 00 03 07 00 00 00 00 00 00 00
    00 00 00 46 6F 6F 3A 00 00 07 00 01 00
);

$result         = _unpack_name($result);
$target         = _unpack_name($target);
is_deeply($result, $target, $caption);




















###############################################################################
#
# Helper functions.
#
###############################################################################

###############################################################################
#
# _unpack_name()
#
# Unpack 1 or more NAME structures into a AoH for easier comparison.
#
sub _unpack_name {

    my $data = shift;
    my @names;

    while ($data) {
        my %name;

        $name{record}       = unpack 'v', substr($data, 0, 2, '');
        $name{length}       = unpack 'v', substr($data, 0, 2, '');
        $name{flags}        = unpack 'v', substr($data, 0, 2, '');
        $name{shortcut}     = unpack 'C', substr($data, 0, 1, '');
        $name{str_len}      = unpack 'C', substr($data, 0, 1, '');
        $name{formula_len}  = unpack 'v', substr($data, 0, 2, '');
        $name{itals}        = unpack 'v', substr($data, 0, 2, '');
        $name{sheet_index}  = unpack 'v', substr($data, 0, 2, '');
        $name{menu_len}     = unpack 'C', substr($data, 0, 1, '');
        $name{desc_len}     = unpack 'C', substr($data, 0, 1, '');
        $name{help_len}     = unpack 'C', substr($data, 0, 1, '');
        $name{status_len}   = unpack 'C', substr($data, 0, 1, '');
        $name{encoding}     = unpack 'C', substr($data, 0, 1, '');


        # Decode the individual flag fields.
        my %flag;
        $flag{hidden}       = $name{flags} & 0x0001;
        $flag{function}     = $name{flags} & 0x0002;
        $flag{vb}           = $name{flags} & 0x0004;
        $flag{macro}        = $name{flags} & 0x0008;
        $flag{complex}      = $name{flags} & 0x0010;
        $flag{builtin}      = $name{flags} & 0x0020;
        $flag{group}        = $name{flags} & 0x0FC0;
        $flag{binary}       = $name{flags} & 0x1000;

        $name{flags} = \%flag;


        # Decode the string part of the NAME structure.
        if ($name{encoding} == 1) {
            # UTF-16 name. Leave in hex.
            $name{string} = uc unpack 'H*', substr($data, 0, 2 * $name{str_len}, '');
        }
        elsif ($flag{'builtin'}) {
            # 1 digit builtin name. Leave in hex.
            $name{string} = uc unpack 'H*', substr($data, 0, $name{str_len}, '');
        }
        else {
            # ASCII name.
            $name{string} = pack 'C*', unpack 'C*', substr($data, 0, $name{str_len}, '');
        }

        # Keep the formula as a hex string.
        $name{formula} = uc unpack 'H*', substr($data, 0, $name{formula_len}, '');

        push @names, \%name;
    }

    return \@names;
}


# Cleanup
$workbook->close();
unlink $test_file;

__END__
