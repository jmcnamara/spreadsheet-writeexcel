package Spreadsheet::Workbook; #Version 0.01

######################################################################
#
# Workbook - A writer class for Excel Workbooks.
#
# Used in conjuction with Spreadsheet::WriteExcel
#
# BETA VERSION OF MULTI-SHEET WORKBOOK
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

use vars qw(@ISA);
@ISA = qw(Spreadsheet::BIFFwriter Exporter);


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
    $self->{_activesheet}    = 0;
    $self->{_firstsheet}     = 0;
    $self->{_fileclosed}     = 0;
    $self->{_biffsize}       = 0;
    $self->{_sheetname}      = "Sheet";
    $self->{worksheets}      = [];

    bless $self, $class;
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
# addworksheet()
#
# Add a new worksheet to the Excel workbook.
# TODO: add accessor for $self->{_sheetname} to mimic international
# versions of Excel
#
# Returns: reference to a worksheet object
#
sub addworksheet {

    my $self      = shift;
    my $name      = $_[0] || "";
    my $index     = @{$self->{worksheets}};
    my $sheetname = $self->{_sheetname};

    if ($name eq "" ) { $name = $sheetname . ($index+1) }

    my @init_data = (
                        $name,
                        $index,
                        \$self->{_activesheet},
                        \$self->{_firstsheet},
                        $self->{_store_in_memory},
                    );
    my $worksheet = Spreadsheet::Worksheet->new(@init_data);
    $self->{worksheets}->[$index] = $worksheet;
    return $worksheet;
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

    if (@{$self->{worksheets}} == 0) { $self->addworksheet() }
    return $self->{worksheets}[0]->write(@_);
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

    if (@{$self->{worksheets}} == 0) { $self->addworksheet() }
    return $self->{worksheets}[0]->write_string(@_);
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

    if (@{$self->{worksheets}} == 0) { $self->addworksheet() }
    return $self->{worksheets}[0]->write_number(@_);
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

    foreach my $sheet (@{$self->{worksheets}}) {
        $offset += $BOF + length($sheet->{name});
    }

    $offset += $EOF;

    foreach my $sheet (@{$self->{worksheets}}) {
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
    foreach my $sheet (@{$self->{worksheets}}) {
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
    foreach my $sheet (@{$self->{worksheets}}) {
        $self->_store_boundsheet($sheet->{name}, $sheet->{_offset});
    }

    # End Workbook globals
    $self->_store_eof();

    # Write Worksheet data if data <~ 7MB
    if ($OLE->set_size($self->{_biffsize})) {
        $OLE->write_header();
        $OLE->write($self->{_data});

        foreach my $sheet (@{$self->{worksheets}}) {
            while (my $tmp = $sheet->get_data()) {
                $OLE->write($tmp);
            }
        }
    }
}


######################################################################
#
# BIFF RECORDS
#


######################################################################
#
# _store_window1()
#
# Write Excel WINDOW1 record.
#
sub _store_window1 {

    my $self      = shift;
    my $name      = 0x003D; # Record identifier
    my $length    = 0x0012; # Number of bytes to follow

    my $xWn       = 0x0000; # Horizontal position of window
    my $yWn       = 0x0069; # Vertical position of window
    my $dxWn      = 0x339F; # Width of window
    my $dyWn      = 0x5D1B; # Height of window

    my $grbit     = 0x0038; # Option flags
    my $ctabsel   = 0x0001; # Number of workbook tabs selected
    my $wTabRatio = 0x0258; # Tab to scrollbar ratio

    my $itabFirst = $self->{_firstsheet};  # 1st displayed worksheet
    my $itabCur   = $self->{_activesheet}; # Selected worksheet

    my $header  = pack("vv",        $name, $length);
    my $data    = pack("vvvvvvvvv", $xWn, $yWn, $dxWn, $dyWn,
                                    $grbit,
                                    $itabCur, $itabFirst,
                                    $ctabsel, $wTabRatio);

    $self->_append($header, $data);
}


######################################################################
#
# _store_font($fontname)
#
# Write Excel FONT record.
#
sub _store_font {

    my $self      = shift;
    my $font      = $_[0];
    my $cch       = length($font);

    my $name      = 0x0031;        # Record identifier
    my $length    = 0x000F + $cch; # Bytes to follow

    my $dyHeight  = 0x00C8; # Height of font (1/20 of a point)
    my $grbit     = 0x0000; # Font attributes
    my $icv       = 0x7FFF; # Index to color palette
    my $bls       = 0x0190; # Bold style
    my $sss       = 0x0000; # Superscript/subscript
    my $uls       = 0x00;   # Underline
    my $bFamily   = 0x00;   # Font family
    my $bCharSet  = 0x00;   # Character set
    my $reserved  = 0x00;   # Reserved

    my $header  = pack("vv",         $name, $length);
    my $data    = pack("vvvvvCCCCC", $dyHeight, $grbit, $icv, $bls,
                                     $sss, $uls, $bFamily, $bCharSet,
                                     $reserved, $cch);

    $self->_append($header, $data, $font);
}


######################################################################
#
# _store_all_fonts()
#
# Write all FONT records.
#
sub _store_all_fonts {

    my $self    = shift;

    $self->_store_font('Arial');
    $self->_store_font('Arial');
    $self->_store_font('Arial');
    $self->_store_font('Arial');
    $self->_store_font('Arial');
}


######################################################################
#
# _store_xf()
#
# Write Excel XF records.
#
sub _store_xf {

    my $self      = shift;
    my $name      = 0x00E0; # Record identifier
    my $length    = 0x0010; # Number of bytes to follow

    my $ifnt      = 0x0000; # Index to FONT record
    my $ifmt      = 0x0000; # Index to FORMAT record
    my $style     = $_[0];  # Style and other options
    my $align     = 0x0020; # Alignment
    my $icv       = 0x20C0; # Color palette and other options
    my $fill      = 0x0000; # Fill and border line style
    my $brd_line  = 0x0000; # Border line style and color
    my $brd_color = 0x0000; # Border color

    my $header  = pack("vv",       $name, $length);
    my $data    = pack("vvvvvvvv", $ifnt, $ifmt, $style, $align,
                                   $icv, $fill,
                                   $brd_line,$brd_color);

    $self->_append($header, $data);
}


######################################################################
#
# _store_all_xfs()
#
# Write all XF records.
#
sub _store_all_xfs {

    my $self    = shift;

    for (0..14) {
        $self->_store_xf(0xFFF5); # Cell XF
    }

    $self->_store_xf(0x0001);     # Style XF
}


######################################################################
#
# _store_style()
#
# Write Excel STYLE records.
#
sub _store_style {

    my $self      = shift;

    my $name      = 0x0093; # Record identifier
    my $length    = 0x0004; # Bytes to follow

    my $ixfe      = 0x0000; # Index to style XF
    my $BuiltIn   = 0x00;   # Built-in style
    my $iLevel    = 0x00;   # Outline style level

    my $header  = pack("vv",  $name, $length);
    my $data    = pack("vCC", $ixfe, $BuiltIn, $iLevel);

    $self->_append($header, $data);
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
# _store_boundsheet()
#
# Writes Excel BOUNDSHEET record.
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


1;


__END__


=head1 NAME

Workbook - A writer class for Excel Workbooks.

=head1 DESCRIPTION

This module is used in conjuction with Spreadsheet::WriteExcel.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

Copyright (c) 2000, John McNamara. All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

