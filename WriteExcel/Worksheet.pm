package Spreadsheet::WriteExcel::Worksheet;

###############################################################################
#
# Worksheet - A writer class for Excel Worksheets.
#
#
# Used in conjuction with Spreadsheet::WriteExcel
#
# Copyright 2000-2001, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

require Exporter;

use strict;
use Carp;
use Spreadsheet::WriteExcel::BIFFwriter;
use Spreadsheet::WriteExcel::Format;



use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::WriteExcel::BIFFwriter);

$VERSION = '0.05';

###############################################################################
#
# new()
#
# Constructor. Creates a new Worksheet object from a BIFFwriter object
#
sub new {

    my $class                = shift;
    my $self                 = Spreadsheet::WriteExcel::BIFFwriter->new();
    my $rowmax               = 65536; # 16384 in Excel 5
    my $colmax               = 256;
    my $strmax               = 255;

    $self->{ name}           = $_[0];
    $self->{_index}          = $_[1];
    $self->{_activesheet}    = $_[2];
    $self->{_firstsheet}     = $_[3];
    $self->{_url_format}     = $_[4];
    $self->{store_in_memory} = $_[5];

    $self->{_filehandle}     = "";
    $self->{_fileclosed}     = 0;
    $self->{_offset}         = 0;
    $self->{_xls_rowmax}     = $rowmax;
    $self->{_xls_colmax}     = $colmax;
    $self->{_xls_strmax}     = $strmax;
    $self->{_dim_rowmin}     = $rowmax +1;
    $self->{_dim_rowmax}     = 0;
    $self->{_dim_colmin}     = $colmax +1;
    $self->{_dim_colmax}     = 0;
    $self->{_colinfo}        = [];
    $self->{_selection}      = [0, 0];


    bless $self, $class;
    $self->_initialize();
    return $self;
}


###############################################################################
#
# _initialize()
#
# If not storing all data in memory open a tmp file for the main
# Worksheet data.
#
sub _initialize {

    my $self    = shift;

    if (not $self->{store_in_memory}) {
        # Open tmp file for storing Worksheet data
        my $fh = IO::File->new_tmpfile();

        if (not defined $fh) {
            croak "Can't open tmp file to store worksheet data.";
        }

        # binmode file whether platform requires it or not
        binmode($fh);

        # Store filehandle
        $self->{_filehandle} = $fh;
    }
}



###############################################################################
#
# _close()
#
# Add data to the beginning of the workbook (note the reverse order)
# and to the end of the workbook.
#
sub _close {

    my $self = shift;

    # Prepend in reverse order!!
    $self->_store_dimensions();


    # Prepend the COLINFO records if they exist
    if (@{$self->{_colinfo}}){
        while (@{$self->{_colinfo}}) {
            my $arrayref = pop @{$self->{_colinfo}};
            $self->_store_colinfo(@$arrayref);
        }
        $self->_store_defcol();
    }

    # Prepend in reverse order!!
    $self->_store_bof(0x0010);

    # Append
    $self->_store_window2();
    $self->_store_selection(@{$self->{_selection}});
    $self->_store_eof();
}


###############################################################################
#
# _append(), overloaded.
#
# Store Worksheet data in memory using the base class _append() or
# to a temporary file, the default.
#
sub _append {

    my $self = shift;

    if ($self->{store_in_memory}) {
        $self->SUPER::_append(@_);
    }
    else {
        my $data    = join('', @_);
        print {$self->{_filehandle}} $data;
        $self->{_datasize} += length($data);
    }
}


###############################################################################
#
# get_data().
#
# Retrieves data from memory in one chunk, or from disk in $buffer
# sized chunks.
#
sub get_data {

    my $self    = shift;
    my $buffer  = 4096;
    my $tmp;

    # Return data stored in memory
    if (defined $self->{_data}) {
        $tmp           = $self->{_data};
        $self->{_data} = undef;
        my $fh         = $self->{_filehandle};
        seek($fh, 0, 0) if not $self->{store_in_memory};
        return $tmp;
    }

    # Return data stored on disk
    if (not $self->{store_in_memory}) {
        return $tmp if read($self->{_filehandle}, $tmp, $buffer);
    }

    # No data to return
    return undef;
}


###############################################################################
#
# activate()
#
# Set this worksheet as the selected worksheet, i.e. the worksheet
# with its tab highlighted.
#
sub activate {

    my $self = shift;

    ${$self->{_activesheet}} = $self->{_index};
}


###############################################################################
#
# set_first_sheet()
#
# Set this worksheet as the first visible sheet. This is necessary
# when there are a large number of worksheets and the activated
# worksheet is not visible on the screen.
#
sub set_first_sheet {

    my $self = shift;

    ${$self->{_firstsheet}} = $self->{_index};
}


###############################################################################
#
# set_column($firstcol, $lastcol, $width, $format, $hidden)
#
# Set the width of a single column or a range of column.
# See also: _store_colinfo
#
sub set_column {

    my $self = shift;

    push @{$self->{_colinfo}}, [ @_ ];
}


###############################################################################
#
# set_col_width()
#
# This is a deprecated alias for set_column().
#
sub set_col_width {

    my $self = shift;

    $self->set_column(@_);
    carp("set_col_width() is deprecated, use set_column() instead") if $^W;
}


###############################################################################
#
# set_selection()
#
# Set which cell or cells are selected in a worksheet: see also the
# sub _store_selection
#
sub set_selection {

    my $self = shift;

    $self->{_selection} = [ @_ ];
}

###############################################################################
#
# _XF()
#
# Returns an index to the XF record in the workbook
#
sub _XF {

    my $self = shift;

    if (ref($self)) {
        return $self->get_xf_index();
    }
    else {
        return 0x0F;
    }
}


###############################################################################
#
# BIFF RECORDS
#

###############################################################################
#
# _store_dimensions()
#
# Writes Excel DIMENSIONS to define the area in which there is data.
#
sub _store_dimensions {

    my $self      = shift;
    my $record    = 0x0000;               # Record identifier
    my $length    = 0x000A;               # Number of bytes to follow
    my $row_min   = $self->{_dim_rowmin}; # First row
    my $row_max   = $self->{_dim_rowmax}; # Last row plus 1
    my $col_min   = $self->{_dim_colmin}; # First column
    my $col_max   = $self->{_dim_colmax}; # Last column plus 1
    my $reserved  = 0x0000;               # Reserved by Excel

    my $header    = pack("vv",    $record, $length);
    my $data      = pack("vvvvv", $row_min, $row_max,
                                  $col_min, $col_max, $reserved);
    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_window2()
#
# Write BIFF record Window2.
#
sub _store_window2 {

    my $self    = shift;
    my $record  = 0x023E;     # Record identifier
    my $length  = 0x000A;     # Number of bytes to follow

    my $grbit   = 0x00B6;     # Option flags
    my $rwTop   = 0x0000;     # Top row visible in window
    my $colLeft = 0x0000;     # Leftmost column visible in window
    my $rgbHdr  = 0x00000000; # Row/column heading and gridline color

    if (${$self->{_activesheet}} == $self->{_index}) {
        $grbit = 0x06B6;
    }

    my $header  = pack("vv",   $record, $length);
    my $data    = pack("vvvV", $grbit, $rwTop, $colLeft, $rgbHdr);

    $self->_append($header, $data);
}


###############################################################################
#
# _store_defcol()
#
# Write BIFF record DEFCOLWIDTH if COLINFO records are in use.
#
sub _store_defcol {

    my $self     = shift;
    my $record   = 0x0055;      # Record identifier
    my $length   = 0x0002;      # Number of bytes to follow

    my $colwidth = 0x0008;      # Default column width

    my $header   = pack("vv", $record, $length);
    my $data     = pack("v",  $colwidth);

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_colinfo($firstcol, $lastcol, $width, $format, $hidden)
#
# Write BIFF record COLINFO to define column widths
#
# Note: The SDK says the record length is 0x0B but Excel writes a 0x0C
# length record.
#
sub _store_colinfo {

    my $self     = shift;
    my $record   = 0x007D;          # Record identifier
    my $length   = 0x000B;          # Number of bytes to follow

    my $colFirst = $_[0] || 0;      # First formatted column
    my $colLast  = $_[1] || 0;      # Last formatted column
    my $coldx    = $_[2] || 8.43;   # Col width, 8.43 is Excel default

    $coldx       += 0.72;           # Fudge. Excel subtracts 0.72 !?
    $coldx       *= 256;            # Convert to units of 1/256 of a char


    my $ixfe     = _XF($_[3]);      # XF
    my $grbit    = $_[4] || 0;      # Option flags
    my $reserved = 0x00;            # Reserved

    my $header   = pack("vv",     $record, $length);
    my $data     = pack("vvvvvC", $colFirst, $colLast, $coldx,
                                  $ixfe, $grbit, $reserved);

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_selection($first_row, $first_col, $last_row, $last_col)
#
# Write BIFF record SELECTION.
#
sub _store_selection {

    my $self     = shift;
    my $record   = 0x001D;                  # Record identifier
    my $length   = 0x000F;                  # Number of bytes to follow

    my $pnn      = 3;                       # Pane position
    my $rwAct    = $_[0];                   # Active row
    my $colAct   = $_[1];                   # Active column
    my $irefAct  = 0;                       # Active cell ref
    my $cref     = 1;                       # Number of refs

    my $rwFirst  = $_[0];                   # First row in reference
    my $colFirst = $_[1];                   # First col in reference
    my $rwLast   = $_[2] || $rwFirst;       # Last row in reference
    my $colLast  = $_[3] || $colFirst;      # Last col in reference

    # Swap last row/col for first row/col as necessary
    if ($rwFirst > $rwLast) {
        ($rwFirst, $rwLast) = ($rwLast, $rwFirst);
    }

    if ($colFirst > $colLast) {
        ($colFirst, $colLast) = ($colLast, $colFirst);
    }


    my $header   = pack("vv",           $record, $length);
    my $data     = pack("CvvvvvvCC",    $pnn, $rwAct, $colAct,
                                        $irefAct, $cref,
                                        $rwFirst, $rwLast,
                                        $colFirst, $colLast);

    $self->_append($header, $data);
}


###############################################################################
#
# write ($row, $col, $token, $format)
#
# Parse $token call appropriate write method. $row and $column are zero
# indexed. $format is optional.
#
# Returns: return value of called subroutine
#
sub write {

    my $self  = shift;
    my $token = $_[2];

    # Match number
    if ($token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/) {
        return $self->write_number(@_);
    }
    # Match http or ftp URL
    elsif ($token =~ m|[fh]tt?p://|) {
        return $self->write_url(@_);
    }
    # Match blank
    elsif ($token eq '') {
        splice @_, 2, 1; # remove the empty string from the parameter list
        return $self->write_blank(@_);
    }
    # Default: match string
    else {
        return $self->write_string(@_);
    }
}


###############################################################################
#
# write_number($row, $col, $num, $format)
#
# Write a double to the specified row and column (zero indexed).
# An integer can be written as a double. Excel will display an
# integer. $format is optional.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub write_number {

    my $self      = shift;
    if (@_ < 3) { return -1 }

    my $record    = 0x0203;                 # Record identifier
    my $length    = 0x000E;                 # Number of bytes to follow

    my $row       = $_[0];                  # Zero indexed row
    my $col       = $_[1];                  # Zero indexed column
    my $num       = $_[2];
    my $xf        = _XF($_[3]);             # The cell format

    if ($row >= $self->{_xls_rowmax}) { return -2 }
    if ($col >= $self->{_xls_colmax}) { return -2 }
    if ($row <  $self->{_dim_rowmin}) { $self->{_dim_rowmin} = $row }
    if ($row >  $self->{_dim_rowmax}) { $self->{_dim_rowmax} = $row }
    if ($col <  $self->{_dim_colmin}) { $self->{_dim_colmin} = $col }
    if ($col >  $self->{_dim_colmax}) { $self->{_dim_colmax} = $col }

    my $header    = pack("vv",  $record, $length);
    my $data      = pack("vvv", $row, $col, $xf);
    my $xl_double = pack("d",   $num);

    if ($self->{_byte_order}) { $xl_double = reverse $xl_double }

    $self->_append($header, $data, $xl_double);

    return 0;
}


###############################################################################
#
# write_string ($row, $col, $string, $format)
#
# Write a string to the specified row and column (zero indexed).
# NOTE: there is an Excel 5 defined limit of 255 characters.
# $format is optional.
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long string truncated to 255 chars
#
sub write_string {

    my $self      = shift;
    if (@_ < 3) { return -1 }

    my $record    = 0x0204;                 # Record identifier
    my $length    = 0x0008 + length($_[2]); # Bytes to follow

    my $row       = $_[0];                  # Zero indexed row
    my $col       = $_[1];                  # Zero indexed column
    my $strlen    = length($_[2]);
    my $str       = $_[2];
    my $xf        = _XF($_[3]);             # The cell format

    my $str_error = 0;

    if ($row >= $self->{_xls_rowmax}) { return -2 }
    if ($col >= $self->{_xls_colmax}) { return -2 }
    if ($row <  $self->{_dim_rowmin}) { $self->{_dim_rowmin} = $row }
    if ($row >  $self->{_dim_rowmax}) { $self->{_dim_rowmax} = $row }
    if ($col <  $self->{_dim_colmin}) { $self->{_dim_colmin} = $col }
    if ($col >  $self->{_dim_colmax}) { $self->{_dim_colmax} = $col }

    if ($strlen > $self->{_xls_strmax}) { # LABEL must be < 255 chars
        $str       = substr($str, 0, $self->{_xls_strmax});
        $length    = 0x0008 + $self->{_xls_strmax};
        $strlen    = $self->{_xls_strmax};
        $str_error = -3;
    }

    my $header    = pack("vv",   $record, $length);
    my $data      = pack("vvvv", $row, $col, $xf, $strlen);

    $self->_append($header, $data, $str);

    return $str_error;
}

###############################################################################
#
# write_blank($row, $col, $format)
#
# Write a blank cell to the specified row and column (zero indexed).
# A blank cell is used to specify formatting without adding a string
# or a number. $format is optional.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub write_blank {

    my $self      = shift;
    if (@_ < 2) { return -1 }

    my $record    = 0x0201;                 # Record identifier
    my $length    = 0x0006;                 # Number of bytes to follow

    my $row       = $_[0];                  # Zero indexed row
    my $col       = $_[1];                  # Zero indexed column
    my $xf        = _XF($_[2]);             # The cell format

    if ($row >= $self->{_xls_rowmax}) { return -2 }
    if ($col >= $self->{_xls_colmax}) { return -2 }
    if ($row <  $self->{_dim_rowmin}) { $self->{_dim_rowmin} = $row }
    if ($row >  $self->{_dim_rowmax}) { $self->{_dim_rowmax} = $row }
    if ($col <  $self->{_dim_colmin}) { $self->{_dim_colmin} = $col }
    if ($col >  $self->{_dim_colmax}) { $self->{_dim_colmax} = $col }

    my $header    = pack("vv",  $record, $length);
    my $data      = pack("vvv", $row, $col, $xf);

    $self->_append($header, $data);

    return 0;
}


###############################################################################
#
# write_url($row, $col, $url, $string, $format )
#
# Write a hyperlink. This is comprised of two elements: the visible label and
# the invisible link. The visible label is the same as the link unless an
# alternative string is specified. The label is written using the
# write_string() method. Therefore the 255 characters string limit applies.
# $string and $format are optional.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long string truncated to 255 chars
#
sub write_url {
    my $self      = shift;
    if (@_ < 3) { return -1 }                    # Check the number of args

    my $record  = 0x01B8;                        # Record identifier
    my $length  = 0x0034 + 2*(1+length($_[2]));  # Bytes to follow

    my $row     = $_[0];                         # Zero indexed row
    my $col     = $_[1];                         # Zero indexed column
    my $url     = $_[2];                         # URL string
    my $str     = $_[3] || $_[2];                # Alternative label
    my $xf      = $_[4] || $self->{_url_format}; # The cell format


    # Write the visible lable using the write_string() method.
    my $str_error = $self->write_string($row, $col, $str, $xf);
    return $str_error if $str_error == -2;


    # Pack the header data
    my $header  = pack("vv",   $record, $length);
    my $data    = pack("vvvv", $row, $row, $col, $col);


    # Pack the undocumented part of the hyperlink stream, 40 bytes.
    my $unknown = "D0C9EA79F9BACE118C8200AA004BA90B02000000";
    $unknown   .= "03000000E0C9EA79F9BACE118C8200AA004BA90B";
    my $stream  = pack("H*", $unknown);


    # Convert URL to a null terminated wchar string
    $url        = join("\0", split('', $url));
    $url        = $url . "\0\0\0";


    # Pack the length of the URL
    my $url_len = pack("V", length($url));


    # Write the packed data
    $self->_append($header, $data);
    $self->_append($stream);
    $self->_append($url_len);
    $self->_append($url);

    return $str_error;

}


###############################################################################
#
# set_row($row, $height, $XF)
#
# This method is used to set the height and XF format for a row.
# Writes the  BIFF record ROW.
#
sub set_row {

    my $self        = shift;
    my $record      = 0x0208;               # Record identifier
    my $length      = 0x0010;               # Number of bytes to follow

    my $rw          = $_[0];                # Row Number
    my $colMic      = 0x0000;               # First defined column
    my $colMac      = 0x0000;               # Last defined column
    my $miyRw;                              # Row height
    my $irwMac      = 0x0000;               # Used by Excel to optimise loading
    my $reserved    = 0x0000;               # Reserved
    my $grbit       = 0x01C0;               # Option flags. (monkey) see $1 do
    my $ixfe        = _XF($_[2]);           # XF index

    # Use set_row($row, undef, $XF) to set XF without setting height
    if (defined ($_[1])) {
        $miyRw = $_[1] *20;
    }
    else {
        $miyRw = 0xff;
    }

    my $header   = pack("vv",       $record, $length);
    my $data     = pack("vvvvvvvv", $rw, $colMic, $colMac, $miyRw,
                                    $irwMac,$reserved, $grbit, $ixfe);

    $self->_append($header, $data);
}


1;


__END__


=head1 NAME

Worksheet - A writer class for Excel Worksheets.

=head1 SYNOPSIS

See the documentation for Spreadsheet::WriteExcel

=head1 DESCRIPTION

This module is used in conjuction with Spreadsheet::WriteExcel.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMI, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
