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
use Test::More tests => 12;
#use Test::More 'no_plan';


###############################################################################
#
# Tests setup
#
my $test_file = 'temp_test_file.xls';
my $workbook;
my $worksheet1;
my $worksheet2;
my $worksheet3;
my $worksheet4;
my $worksheet5;
my $target;
my $result;
my $caption;
my $area;


###############################################################################
#
# Tests 1, 2. Test print_area() for a simple range.
#
$workbook                   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1                 = $workbook->add_worksheet();
$workbook->{_using_tmpfile} = 0;
$area                       = 'A1:B12';

$worksheet1->print_area($area);

# Test the EXTERNSHEET record.
$workbook->_calculate_extern_sizes();
$workbook->_store_externsheet();

$target     = pack 'H*', join '', qw(
    17 00 08 00 01 00 00 00 00 00 00 00
);

$caption    = " \tExternsheet";
$result     = _unpack_externsheet($workbook->{_data});
$target     = _unpack_externsheet($target);
is_deeply($result, $target, $caption);


# Test the NAME record.
$workbook->{_data} = '';
$workbook->_store_names();

$target     = pack 'H*', join '', qw(
    18 00 1B 00 20 00 00 01 0B 00 00 00 01 00 00 00
    00 00 00 06 3B 00 00 00 00 0B 00 00 00 01 00
);

$caption    = " \t+ Name      ( Sheet1!$area )";
$result     = _unpack_name($workbook->{_data});
$target     = _unpack_name($target);
is_deeply($result, $target, $caption);

$workbook->close();


###############################################################################
#
# Tests 3, 4. Test print_area() for a single row.
#
$workbook                   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1                 = $workbook->add_worksheet();
$workbook->{_using_tmpfile} = 0;
$area                       = 'A1:IV1';

$worksheet1->print_area($area);

# Test the EXTERNSHEET record.
$workbook->_calculate_extern_sizes();
$workbook->_store_externsheet();

$target     = pack 'H*', join '', qw(
    17 00 08 00 01 00 00 00 00 00 00 00
);

$caption    = " \tExternsheet";
$result     = _unpack_externsheet($workbook->{_data});
$target     = _unpack_externsheet($target);
is_deeply($result, $target, $caption);


# Test the NAME record.
$workbook->{_data} = '';
$workbook->_store_names();

$target     = pack 'H*', join '', qw(
    18 00 1B 00 20 00 00 01 0B 00 00 00 01 00 00 00
    00 00 00 06 3B 00 00 00 00 00 00 00 00 FF 00
);

$caption    = " \t+ Name      ( Sheet1!$area )";
$result     = _unpack_name($workbook->{_data});
$target     = _unpack_name($target);
is_deeply($result, $target, $caption);

$workbook->close();


###############################################################################
#
# Tests 5, 6. Test print_area() for a single column.
#
$workbook                   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1                 = $workbook->add_worksheet();
$workbook->{_using_tmpfile} = 0;
$area                       = 'A1:A65536';

$worksheet1->print_area($area);

# Test the EXTERNSHEET record.
$workbook->_calculate_extern_sizes();
$workbook->_store_externsheet();

$target     = pack 'H*', join '', qw(
    17 00 08 00 01 00 00 00 00 00 00 00
);

$caption    = " \tExternsheet";
$result     = _unpack_externsheet($workbook->{_data});
$target     = _unpack_externsheet($target);
is_deeply($result, $target, $caption);


# Test the NAME record.
$workbook->{_data} = '';
$workbook->_store_names();

$target     = pack 'H*', join '', qw(
    18 00 1B 00 20 00 00 01 0B 00 00 00 01 00 00 00
    00 00 00 06 3B 00 00 00 00 FF FF 00 00 00 00
);

$caption    = " \t+ Name      ( Sheet1!$area )";
$result     = _unpack_name($workbook->{_data});
$target     = _unpack_name($target);
is_deeply($result, $target, $caption);

$workbook->close();


###############################################################################
#
# Tests 7, 8. Test print_area() for multiple columns.
#
$workbook                   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1                 = $workbook->add_worksheet();
$workbook->{_using_tmpfile} = 0;
$area                       = 'A:H';

$worksheet1->print_area($area);

# Test the EXTERNSHEET record.
$workbook->_calculate_extern_sizes();
$workbook->_store_externsheet();

$target     = pack 'H*', join '', qw(
    17 00 08 00 01 00 00 00 00 00 00 00
);

$caption    = " \tExternsheet";
$result     = _unpack_externsheet($workbook->{_data});
$target     = _unpack_externsheet($target);
is_deeply($result, $target, $caption);


# Test the NAME record.
$workbook->{_data} = '';
$workbook->_store_names();

$target     = pack 'H*', join '', qw(
    18 00 1B 00 20 00 00 01 0B 00 00 00 01 00 00 00
    00 00 00 06 3B 00 00 00 00 FF FF 00 00 07 00
);

$caption    = " \t+ Name      ( Sheet1!$area )";
$result     = _unpack_name($workbook->{_data});
$target     = _unpack_name($target);
is_deeply($result, $target, $caption);

$workbook->close();


###############################################################################
#
# Tests 9, 10. Test ranges on multiple sheets.
#
$workbook                   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1                 = $workbook->add_worksheet();
$worksheet2                 = $workbook->add_worksheet();
$worksheet3                 = $workbook->add_worksheet();
$workbook->{_using_tmpfile} = 0;


$worksheet1->print_area('A1:B2');
$worksheet2->print_area('D4:E5');
$worksheet3->print_area('G7:H8');

# Test the EXTERNSHEET record.
$workbook->_calculate_extern_sizes();
$workbook->_store_externsheet();

$target     = pack 'H*', join '', qw(
    17 00 14 00 03 00 00 00 00 00 00 00 00 00 01 00
    01 00 00 00 02 00 02 00
);

$caption    = " \tExternsheet";
$result     = _unpack_externsheet($workbook->{_data});
$target     = _unpack_externsheet($target);
is_deeply($result, $target, $caption);


# Test the NAME record.
$workbook->{_data} = '';
$workbook->_store_names();

$target     = pack 'H*', join '', qw(
    18 00 1B 00 20 00 00 01 0B 00 00 00 01 00 00 00
    00 00 00 06 3B 00 00 00 00 01 00 00 00 01 00

    18 00 1B 00 20 00 00 01 0B 00 00 00 02 00 00 00
    00 00 00 06 3B 01 00 03 00 04 00 03 00 04 00

    18 00 1B 00 20 00 00 01 0B 00 00 00 03 00 00 00
    00 00 00 06 3B 02 00 06 00 07 00 06 00 07 00
);

$caption    = " \t+ Name      ( Sheet1!A1:B2, Sheet2!D4:E5, Sheet3!G7:H8 )";
$result     = _unpack_name($workbook->{_data});
$target     = _unpack_name($target);
is_deeply($result, $target, $caption);

$workbook->close();


###############################################################################
#
# Tests 11, 12. Test ranges on multiple sheets, with sheets spaced out.
#
$workbook                   = Spreadsheet::WriteExcel->new($test_file);
$worksheet1                 = $workbook->add_worksheet();
$worksheet2                 = $workbook->add_worksheet();
$worksheet3                 = $workbook->add_worksheet();
$worksheet4                 = $workbook->add_worksheet();
$worksheet5                 = $workbook->add_worksheet();
$workbook->{_using_tmpfile} = 0;


$worksheet1->print_area('A1:B2');
$worksheet3->print_area('D4:E5');
$worksheet5->print_area('G7:H8');

# Test the EXTERNSHEET record.
$workbook->_calculate_extern_sizes();
$workbook->_store_externsheet();

$target     = pack 'H*', join '', qw(
    17 00 14 00 03 00 00 00 00 00 00 00 00 00 02 00
    02 00 00 00 04 00 04 00
);

$caption    = " \tExternsheet";
$result     = _unpack_externsheet($workbook->{_data});
$target     = _unpack_externsheet($target);
is_deeply($result, $target, $caption);


# Test the NAME record.
$workbook->{_data} = '';
$workbook->_store_names();

$target     = pack 'H*', join '', qw(
    18 00 1B 00 20 00 00 01 0B 00 00 00 01 00 00 00
    00 00 00 06 3B 00 00 00 00 01 00 00 00 01 00

    18 00 1B 00 20 00 00 01 0B 00 00 00 03 00 00 00
    00 00 00 06 3B 01 00 03 00 04 00 03 00 04 00

    18 00 1B 00 20 00 00 01 0B 00 00 00 05 00 00 00
    00 00 00 06 3B 02 00 06 00 07 00 06 00 07 00
);

$caption    = " \t+ Name      ( Sheet1!A1:B2, Sheet3!D4:E5, Sheet5!G7:H8 )";
$result     = _unpack_name($workbook->{_data});
$target     = _unpack_name($target);
is_deeply($result, $target, $caption);

$workbook->close();


###############################################################################
#
# Helper functions.
#
###############################################################################


###############################################################################
#
# _unpack_externsheet()
#
# Unpack the EXTERNSHEET recordfor easier comparison.
#
sub _unpack_externsheet {

    my $data = shift;
    my %externsheet;

    $externsheet{record}       = unpack 'v', substr($data, 0, 2, '');
    $externsheet{length}       = unpack 'v', substr($data, 0, 2, '');
    $externsheet{count}        = unpack 'v', substr($data, 0, 2, '');
    $externsheet{array}        = [];

    for (1 .. $externsheet{count}) {
        push @{$externsheet{array}},
              [unpack 'vvv', substr($data, 0, 6, '')];
    }

    return \%externsheet;
}


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
unlink $test_file;

__END__
