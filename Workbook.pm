package Spreadsheet::Workbook;

######################################################################
#
# Workbook - A writer class for Excel Workbooks.
#
#
# Used in conjuction with Spreadsheet::WriteExcel
#
# Copyright 2000, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

require Exporter;

use strict;
use Spreadsheet::BIFFwriter;
use Spreadsheet::Worksheet;
use Spreadsheet::OLEwriter;
use Spreadsheet::Format;

use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::BIFFwriter Exporter);

$VERSION = '0.03';

######################################################################
#
# new()
#
# Constructor. Creates a new Workbook object from a BIFFwriter object.
#
sub new {

    my $class                = shift;
    my $self                 = Spreadsheet::BIFFwriter->new();
    my $filename             = $_[0] || '';

    $self->{_store_in_memory}= $_[1] || 0;
    $self->{_OLEwriter}      = Spreadsheet::OLEwriter->new($filename);
    $self->{_active_sheet}   = 0;
    $self->{_first_sheet}    = 0;
    $self->{_fileclosed}     = 0;
    $self->{_biffsize}       = 0;
    $self->{_sheetname}      = "Sheet";
    $self->{_tmp_worksheet}  = Spreadsheet::Worksheet->new('', 0, 0);
    $self->{_tmp_format}     = Spreadsheet::Format->new();
    $self->{_worksheets}     = [];
    $self->{_formats}        = [];

    bless $self, $class;
    #$self->addformat(); # Used to write the default XFs
    return $self;
}


######################################################################
#
# close()
#
# Calls finalization methods and explicitly close the OLEwriter file
# handle.
#
sub close {

    my $self = shift;

    return if $self->{_fileclosed}; # Prevent calling close() twice

    $self->_store_workbook();
    $self->{_OLEwriter}->close();
    $self->{_fileclosed} = 1;
}


######################################################################
#
# DESTROY()
#
# Close the workbook if it hasn't already been explicitly closed.
#
sub DESTROY {

    my $self = shift;

    $self->close() if not $self->{_fileclosed};
}


######################################################################
#
# worksheets()
#
# An accessor for the _worksheets[] array
#
# Returns: an array reference
#
sub worksheets {

    my $self = shift;

    return $self->{_worksheets};
}


######################################################################
#
# addworksheet()
#
# Add a new worksheet to the Excel workbook.
# TODO: add accessor for $self->{_sheetname} to mimic international
# versions of Excel.
#
# Returns: reference to a worksheet object
#
sub addworksheet {

    my $self      = shift;
    my $name      = $_[0] || "";
    my $index     = @{$self->{_worksheets}};
    my $sheetname = $self->{_sheetname};

    if ($name eq "" ) { $name = $sheetname . ($index+1) }

    my @init_data = (
                        $name,
                        $index,
                        $self->{_store_in_memory},
                    );

    my $worksheet = Spreadsheet::Worksheet->new(@init_data);
    $self->{_worksheets}->[$index] = $worksheet;
    return $worksheet;
}

######################################################################
#
# addformat()
#
# Add a new format to the Excel workbook. This adds an XF record and
# a FONT record. TODO: add a FORMAT record.
#
sub addformat {

    my $self      = shift;

    my $format = Spreadsheet::Format->new();
    push @{$self->{_formats}}, $format;
    return $format;
}



######################################################################
#
# write()
#
# Calls write method on first worksheet for backward compatibility.
# Adds first worksheet as necessary.
#
# Returns: return value of the worksheet->write() method
#
sub write {

    my $self    = shift;

    if (@{$self->{_worksheets}} == 0) { $self->addworksheet() }
    return $self->{_worksheets}[0]->write(@_);
}


######################################################################
#
# write_string()
#
# Calls write_string method on first worksheet for backward
# compatibility. Adds first worksheet as necessary.
#
# Returns: return value of the worksheet->write_string() method
#
sub write_string {

    my $self    = shift;

    if (@{$self->{_worksheets}} == 0) { $self->addworksheet() }
    return $self->{_worksheets}[0]->write_string(@_);
}


######################################################################
#
# write_number()
#
# Calls write_number method on first worksheet for backward
# compatibility. Adds first worksheet as necessary.
#
# Returns: return value of the worksheet->write_number() method
#
sub write_number {

    my $self    = shift;

    if (@{$self->{_worksheets}} == 0) { $self->addworksheet() }
    return $self->{_worksheets}[0]->write_number(@_);
}

######################################################################
#
# _calc_sheet_offsets()
#
# Calculate offsets for Worksheet BOF records.
#
sub _calc_sheet_offsets {

    my $self    = shift;
    my $BOF     = 11;
    my $EOF     = 4;
    my $offset  = $self->{_datasize};

    foreach my $sheet (@{$self->{_worksheets}}) {
        $offset += $BOF + length($sheet->{name});
    }

    $offset += $EOF;

    foreach my $sheet (@{$self->{_worksheets}}) {
        $sheet->{_offset} = $offset;
        $offset += $sheet->{_datasize};
    }

    $self->{_biffsize} = $offset;
}


######################################################################
#
# _store_workbook()
#
# Assemble worksheets into a workbook and send the BIFF data to an OLE
# storage.
#
sub _store_workbook {

    my $self = shift;
    my $OLE  = $self->{_OLEwriter};

    # Call the finalization methods for each worksheet
    foreach my $sheet (@{$self->{_worksheets}}) {
        $sheet->_close();
    }

    # Add Workbook globals
    $self->_store_bof(0x0005);
    $self->_store_window1();
    $self->_store_all_fonts();
    $self->_store_all_xfs();
    $self->_store_all_styles();
    $self->_calc_sheet_offsets();

    # Add BOUNDSHEET records
    foreach my $sheet (@{$self->{_worksheets}}) {
        $self->_store_boundsheet($sheet->{name}, $sheet->{_offset});
    }

    # End Workbook globals
    $self->_store_eof();

    # Write Worksheet data if data <~ 7MB
    if ($OLE->set_size($self->{_biffsize})) {
        $OLE->write_header();
        $OLE->write($self->{_data});

        foreach my $sheet (@{$self->{_worksheets}}) {
            while (my $tmp = $sheet->get_data()) {
                $OLE->write($tmp);
            }
        }
    }
}


######################################################################
#
# _store_all_fonts()
#
# Store the Excel FONT records. TODO: At present a FONT record is
# created for each XF record. Duplicates should be removed.
sub _store_all_fonts {

    my $self   = shift;

    # _tmp_format is added by new() and not by the user.
    # We use  this to write the default XF's
    my $format = $self->{_tmp_format};
    my $font   = $format->get_font();

    for (1..4){
        $self->_append($font);
    }

    # User defined fonts
    foreach $format (@{$self->{_formats}}) {
        $font = $format->get_font();
        $self->_append($font);
    }
}


######################################################################
#
# _store_all_xfs()
#
# Write all XF records.
#
sub _store_all_xfs {

    my $self    = shift;

    # _tmp_format is added by new() and not by the user.
    # We use  this to write the default XF's
    my $format = $self->{_tmp_format};
    my $xf;

    # Set the font attribute to the default value
    $format->{_font_index} = 0;

    for (0..14) {
        $xf = $format->get_xf(0xFFF5); # Style XF
        $self->_append($xf);
    }

    $xf = $format->get_xf(0x0001);     # Cell XF
    $self->_append($xf);


    # User defined formats
    foreach $format (@{$self->{_formats}}) {
        $xf = $format->get_xf(0x0001);
        $self->_append($xf);
    }
}


######################################################################
#
# _store_all_styles()
#
# Write all STYLE records.
#
sub _store_all_styles {

    my $self    = shift;

    $self->_store_style();
}


######################################################################
#
# BIFF RECORDS
#


######################################################################
#
# _store_window1()
#
# Write Excel BIFF WINDOW1 record.
#
sub _store_window1 {

    my $self      = shift;
    my $name      = 0x003D;  # Record identifier
    my $length    = 0x0012;  # Number of bytes to follow

    my $xWn       = 0x0000;  # Horizontal position of window
    my $yWn       = 0x0000;  # Vertical position of window
    my $dxWn      = 0x25BC;  # Width of window
    my $dyWn      = 0x1572;  # Height of window

    my $grbit     = 0x0038;  # Option flags
    my $ctabsel   = 0x0001;  # Number of workbook tabs selected
    my $wTabRatio = 0x0258;  # Tab to scrollbar ratio


    my $worksheet = $self->{_tmp_worksheet};
    my $first     = $worksheet->get_first_sheet();
    my $active    = $worksheet->get_active_sheet();

    my $itabFirst = $first;  # 1st displayed worksheet
    my $itabCur   = $active; # Selected worksheet

    my $header    = pack("vv",        $name, $length);
    my $data      = pack("vvvvvvvvv", $xWn, $yWn, $dxWn, $dyWn,
                                      $grbit,
                                      $itabCur, $itabFirst,
                                      $ctabsel, $wTabRatio);

    $self->_append($header, $data);
}


######################################################################
#
# _store_boundsheet()
#
# Writes Excel BIFF BOUNDSHEET record.
#
sub _store_boundsheet {

    my $self      = shift;

    my $name      = 0x0085;               # Record identifier
    my $length    = 0x07 + length($_[0]); # Number of bytes to follow

    my $sheetname = $_[0];                # Worksheet name
    my $offset    = $_[1];                # Location of worksheet BOF
    my $grbit     = 0x0000;               # Sheet identifier
    my $cch       = length($sheetname);   # Length of sheet name

    my $header    = pack("vv",  $name, $length);
    my $data      = pack("VvC", $offset, $grbit, $cch);

    $self->_append($header, $data, $sheetname)
}

######################################################################
#
# _store_style()
#
# Write Excel BIFF STYLE records.
#
sub _store_style {

    my $self      = shift;

    my $name      = 0x0093; # Record identifier
    my $length    = 0x0004; # Bytes to follow

    my $ixfe      = 0x0000; # Index to style XF
    my $BuiltIn   = 0x00;   # Built-in style
    my $iLevel    = 0x00;   # Outline style level

    my $header    = pack("vv",  $name, $length);
    my $data      = pack("vCC", $ixfe, $BuiltIn, $iLevel);

    $self->_append($header, $data);
}


1;


__END__


=head1 NAME

Workbook - A writer class for Excel Workbooks.

=head1 SYNOPSIS

See the documentation for Spreadsheet::WriteExcel

=head1 DESCRIPTION

This module is used in conjuction with Spreadsheet::WriteExcel.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

Copyright (c) 2000, John McNamara. All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

