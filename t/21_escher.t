#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Tests for the internal methods used to write the records in an Escher drawing
# object such as images, comments and filters.
#
# reverse('©'), September 2005, John McNamara, jmcnamara@cpan.org
#


use strict;

use Spreadsheet::WriteExcel;
use Test::More tests => 42;


my @tests = (
                [   'DggContainer',                 # Caption
                    0xF000,                         # Type
                    15,                             # Version
                    0,                              # Instance
                    '',                             # Data
                    82,                             # Length
                    '0F 00 00 F0 52 00 00 00',      # Target
                ],

                [   'DgContainer',                  # Caption
                    0xF002,                         # Type
                    15,                             # Version
                    0,                              # Instance
                    '',                             # Data
                    328,                            # Length
                    '0F 00 02 F0 48 01 00 00',      # Target
                ],

                [   'SpgrContainer',                # Caption
                    0xF003,                         # Type
                    15,                             # Version
                    0,                              # Instance
                    '',                             # Data
                    304,                            # Length
                    '0F 00 03 F0 30 01 00 00',      # Target
                ],

                [   'SpContainer',                  # Caption
                    0xF004,                         # Type
                    15,                             # Version
                    0,                              # Instance
                    '',                             # Data
                    40,                             # Length
                    '0F 00 04 F0 28 00 00 00',      # Target
                ],

                [   'Dgg',                          # Caption
                    0xF006,                         # Type
                    0,                              # Version
                    0,                              # Instance
                    '02 04 00 00 02 00 00 00 ' .    # Data
                    '02 00 00 00 01 00 00 00 ' .
                    '01 00 00 00 02 00 00 00',
                    undef,                          # Length
                    '00 00 06 F0 18 00 00 00 ' .    # Target
                    '02 04 00 00 02 00 00 00 ' .
                    '02 00 00 00 01 00 00 00 ' .
                    '01 00 00 00 02 00 00 00',
                ],

                [   'Dg',                           # Caption
                    0xF008,                         # Type
                    0,                              # Version
                    1,                              # Instance
                    '03 00 00 00 02 04 00 00',      # Data
                    undef,                          # Length
                    '10 00 08 F0 08 00 00 00 ' .    # Target
                    '03 00 00 00 02 04 00 00',
                ],

                [   'Spgr',                         # Caption
                    0xF009,                         # Type
                    1,                              # Version
                    0,                              # Instance
                    '00 0E 00 0E 40 41 00 00 ' .    # Data
                    '00 0E 00 0E 40 41 00 00',
                    undef,                          # Length
                    '01 00 09 F0 10 00 00 00 ' .    # Target
                    '00 0E 00 0E 40 41 00 00 ' .
                    '00 0E 00 0E 40 41 00 00',
                ],

                [   'ClientTextbox',                # Caption
                    0xF00D,                         # Type
                    0,                              # Version
                    0,                              # Instance
                    '',                             # Data
                    undef,                          # Length
                    '00 00 0D F0 00 00 00 00',      # Target
                ],

                [   'ClientAnchor',                 # Caption
                    0xF010,                         # Type
                    0,                              # Version
                    0,                              # Instance
                    '03 00 01 00 F0 00 01 00 ' .    # Data
                    '69 00 03 00 F0 00 05 00 ' .
                    'C4 00',
                    undef,                          # Length
                    '00 00 10 F0 12 00 00 00 ' .    # Target
                    '03 00 01 00 F0 00 01 00 ' .
                    '69 00 03 00 F0 00 05 00 ' .
                    'C4 00',
                ],

                [   'ClientData',                   # Caption
                    0xF011,                         # Type
                    0,                              # Version
                    0,                              # Instance
                    '',                             # Data
                    undef,                          # Length
                    '00 00 11 F0 00 00 00 00',      # Target
                ],

                [   'SplitMenuColors',              # Caption
                    0xF11E,                         # Type
                    0,                              # Version
                    4,                              # Instance
                    '0D 00 00 08 0C 00 00 08 ' .    # Data
                    '17 00 00 08 F7 00 00 10',
                    undef,                          # Length
                    '40 00 1E F1 10 00 00 00 ' .    # Target
                    '0D 00 00 08 0C 00 00 08 ' .
                    '17 00 00 08 F7 00 00 10',
                ],

                [   'BstoreContainer',              # Caption
                    0xF001,                         # Type
                    15,                             # Version
                    1,                              # Instance
                    '',                             # Data
                    163,                            # Length
                    '1F 00 01 F0 A3 00 00 00',      # Target
                ],
            );



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
my @data;
my $range;


###############################################################################
#
# Tests for the generic method.
#
for my $aref (@tests) {

    my @data    = @$aref;
    my $caption = shift @data;
    my $target  = pop   @data;

    $data[3]    =~ s/ //g;
    $data[3]    =  pack "H*", $data[3];

    $caption    = sprintf " \t_add_mso_generic(): (0x%04X) %s", $data[0], $caption;

    $result     = unpack_record($worksheet->_add_mso_generic(@data));
    is($result, $target, $caption);
}


###############################################################################
#
# Test for _store_mso_dg_container.
#
$caption = sprintf " \t_store_mso_dgg_container()";
@data    = ();
$target  = join ' ', qw(
                        0F 00 00 F0 52 00 00 00
                       );


$workbook->{_mso_size} = 94;
$result  = unpack_record($workbook->_store_mso_dgg_container(@data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_dgg.
#
$caption = sprintf " \t_store_mso_dgg()";
@data    = (1026, 2, 2, 1, [[1,2]]);
$target  = join ' ', qw(
                        00 00 06 F0
                        18 00 00 00 02 04 00 00 02 00 00 00 02 00 00 00
                        01 00 00 00 01 00 00 00 02 00 00 00
                       );

$result  = unpack_record($workbook->_store_mso_dgg(@data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_opt.
#
$caption = sprintf " \t_store_mso_opt()";
@data    = ();
$target  = join ' ', qw(
                        33 00 0B F0
                        12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
                        C0 01 40 00 00 08
                       );

$result  = unpack_record($workbook->_store_mso_opt(@data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_split_menu_colors.
#
$caption = sprintf " \t_store_mso_split_menu_colors()";
@data    = ();
$target  = join ' ', qw(
                        40 00 1E F1 10 00 00 00 0D 00
                        00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
                       );

$result  = unpack_record($workbook->_store_mso_split_menu_colors(@data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_dg_container.
#
$caption = sprintf " \t_store_mso_dg_container()";
@data    = (0xC8);
$target  = join ' ', qw(
                        0F 00 02 F0 C8 00 00 00
                       );

$result  = unpack_record($worksheet->_store_mso_dg_container(@data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_dg.
#
$caption = sprintf " \t_store_mso_dg()";
@data    = (1, 2, 1025);
$target  = join ' ', qw(
                        10 00 08 F0
                        08 00 00 00 02 00 00 00 01 04 00 00
                       );

$result  = unpack_record($worksheet->_store_mso_dg(@data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_spgr_container.
#
$caption = sprintf " \t_store_mso_spgr_container()";
@data    = (0xB0);
$target  = join ' ', qw(
                        0F 00 03 F0 B0 00 00 00
                       );

$result  = unpack_record($worksheet->_store_mso_spgr_container(@data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_sp_container.
#
$caption = sprintf " \t_store_mso_sp_container()";
@data    = (0x28);
$target  = join ' ', qw(
                        0F 00 04 F0 28 00 00 00
                       );

$result  = unpack_record($worksheet->_store_mso_sp_container(@data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_sp.
#
$caption = sprintf " \t_store_mso_sp()";
@data    = (0, 1024, 0x0005);
$target  = join ' ', qw(
                        02 00 0A F0 08 00 00 00 00 04 00 00 05 00 00 00
                       );

$result  = unpack_record($worksheet->_store_mso_sp(@data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_sp.
#
$caption = sprintf " \t_store_mso_sp()";
@data    = (202, 1025, 0x0A00);
$target  = join ' ', qw(
                        A2 0C 0A F0 08 00 00 00 01 04 00 00 00 0A 00 00
                       );

$result  = unpack_record($worksheet->_store_mso_sp(@data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_opt_comment.
#
$caption = sprintf " \t_store_mso_opt_comment()";
@data    = (0x80);
$target  = join ' ', qw(
                        93 00 0B F0 36 00 00 00
                        80 00 00 00 00 00 BF 00 08 00 08 00
                        58 01 00 00 00 00 81 01 50 00 00 08 83 01 50 00
                        00 08 BF 01 10 00 11 00 01 02 00 00 00 00 3F 02
                        03 00 03 00 BF 03 02 00 0A 00
                       );

$result  = unpack_record($worksheet->_store_mso_opt_comment(@data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor.
#
$range   = 'A1';
$caption = sprintf " \t_store_mso_client_anchor(%s)", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test');
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00 00 00 03 00 01 00 F0 00 00 00
                        1E 00 03 00 F0 00 04 00 78 00
                       );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor.
#
$range   = 'A2';
$caption = sprintf " \t_store_mso_client_anchor(%s)", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test');
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00 00 00 03 00 01 00 F0 00 00 00
                        69 00 03 00 F0 00 04 00 C4 00
                       );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor.
#
$range   = 'A3';
$caption = sprintf " \t_store_mso_client_anchor(%s)", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test');
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00 00 00 03 00 01 00 F0 00 01 00
                        69 00 03 00 F0 00 05 00 C4 00
                       );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor.
#
$range   = 'A65534';
$caption = sprintf " \t_store_mso_client_anchor(%s)", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test');
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00 00 00 03 00 01 00 F0 00 F9 FF
                        3C 00 03 00 F0 00 FD FF 97 00
                       );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor.
#
$range   = 'A65535';
$caption = sprintf " \t_store_mso_client_anchor(%s)", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test');
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00 00 00 03 00 01 00 F0 00 FA FF
                        3C 00 03 00 F0 00 FE FF 97 00
                       );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor.
#
$range   = 'A65536';
$caption = sprintf " \t_store_mso_client_anchor(%s)", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test');
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00 00 00 03 00 01 00 F0 00 FB FF
                        1E 00 03 00 F0 00 FF FF 78 00
                       );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor.
#
$range   = 'IT3';
$caption = sprintf " \t_store_mso_client_anchor(%s)", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test');
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00 00 00 03 00 FA 00 10 03 01 00
                        69 00 FC 00 10 03 05 00 C4 00
                       );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor.
#
$range   = 'IU3';
$caption = sprintf " \t_store_mso_client_anchor(%s)", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test');
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00 00 00 03 00 FB 00 10 03 01 00
                        69 00 FD 00 10 03 05 00 C4 00
                       );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor.
#
$range   = 'IV3';
$caption = sprintf " \t_store_mso_client_anchor(%s)", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test');
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00 00 00 03 00 FC 00 10 03 01 00
                        69 00 FE 00 10 03 05 00 C4 00
                       );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor where comment offsets have changed.
#
$range   = 'A3';

$caption = sprintf " \t_store_mso_client_anchor(%s). Cell offsets changed.", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test', x_offset => 18, y_offset => 9);
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00
                        00 00 03 00 01 00 20 01 01 00 88 00 03 00 20 01
                        05 00 E2 00
                        );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor where comment dimensions have changed.
#
$range   = 'A3';

$caption = sprintf " \t_store_mso_client_anchor(%s). Dimensions changed.", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test', x_scale => 3, y_scale => 2);
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00 00 00 03 00
                        01 00 F0 00 01 00 69 00 07 00 F0 00 0A 00 1E 00

                        );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor where comment dimensions have changed.
#
$range   = 'A3';

$caption = sprintf " \t_store_mso_client_anchor(%s). Dimensions changed.", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test', width => 384, height => 148);
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00 00 00 03 00
                        01 00 F0 00 01 00 69 00 07 00 F0 00 0A 00 1E 00

                        );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor where column widths have changed.
#
$range   = 'F3';
$worksheet->set_column('G:G', 20);

$caption = sprintf " \t_store_mso_client_anchor(%s). Col width changed.", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test');
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00
                        00 00 03 00 06 00 6A 00 01 00 69 00 06 00 F2 03
                        05 00 C4 00
                        );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor where column widths have changed.
#
$range   = 'K3';
$worksheet->set_column('L:O', 4);

$caption = sprintf " \t_store_mso_client_anchor(%s). Col width changed.", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test');
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00
                        00 00 03 00 0B 00 D1 01 01 00 69 00 0F 00 B0 00
                        05 00 C4 00
                        );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor where row height have changed.
#
$range   = 'A6';
$worksheet->set_row(5,   6);
$worksheet->set_row(6,   6);
$worksheet->set_row(7,   6);
$worksheet->set_row(8,   6);

$caption = sprintf " \t_store_mso_client_anchor(%s). Row height changed.", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test');
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00
                        00 00 03 00 01 00 F0 00 04 00 69 00 03 00 F0 00
                        0A 00 E2 00
                        );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_anchor where row height have changed.
#
$range   = 'A15';
$worksheet->set_row(14, 60);

$caption = sprintf " \t_store_mso_client_anchor(%s). Row height changed.", $range;
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test');
@data    = @{$data[-1]};
$target  = join ' ', qw(
                        00 00 10 F0 12 00
                        00 00 03 00 01 00 F0 00 0D 00 69 00 03 00 F0 00
                        0E 00 CD 00
                        );

$result  = unpack_record($worksheet->_store_mso_client_anchor(3, @data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_data.
#
$caption = sprintf " \t_store_mso_client_data()";
@data    = ();
$target  = join ' ', qw(
                        00 00 11 F0 00 00 00 00
                       );

$result  = unpack_record($worksheet->_store_mso_client_data(@data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_obj_comment.
#
$caption = sprintf " \t_store_obj_comment()";
@data    = (0x01);
$target  = join ' ', qw(
                        5D 00 34 00 15 00 12 00 19 00 01 00 11 40 00 00
                        00 00 00 00 00 00 00 00 00 00 0D 00 16 00 00 00
                        00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00
                        00 00 00 00 00 00 00 00
                       );

$result  = unpack_record($worksheet->_store_obj_comment(@data));

is($result, $target, $caption);


###############################################################################
#
# Test for _store_mso_client_text_box.
#
$caption = sprintf " \t_store_mso_client_text_box()";
@data    = ();
$target  = join ' ', qw(
                        00 00 0D F0 00 00 00 00
                       );

$result  = unpack_record($worksheet->_store_mso_client_text_box(@data));

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

