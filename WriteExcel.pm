package Spreadsheet::WriteExcel;

######################################################################
#
# WriteExcel - Write text and numbers to minimal Excel binary file.
#
# Copyright 1999-2000, John McNamara, john.exeng@abanet.it
#
# Documentation after __END__
#

require Exporter;

use strict;
use Carp;
use FileHandle;

use vars qw($VERSION @ISA);

@ISA = qw(Exporter);
$VERSION = '0.08'; # 16 Jan 2000, Berryman


######################################################################
#
# Constructor
#

sub new {
    my $class = $_[0];
    my $self  = {
                    _xlsfilename   => $_[1] || "",
                    _filehandle    => "",
                    _byte_order    => "",
                    _maxrow        => 65536,
                    _maxcol        => 256,
                    _maxstr        => 255,
                 };

    bless $self, $class;
    $self->_initialize();
    return $self;
}


######################################################################
#
# Initialization routines
#

sub _initialize {
    my $self    = shift;
    my $xlsfile = $self->{_xlsfilename};

    # Check for filename
    if ($xlsfile eq "") {
        croak('Filename required in WriteExcel("Filename")');
    }

    # Open file for writing and reading (to read size)
    my $fh = FileHandle->new("+>$xlsfile");
    if (not defined $fh) {
        croak "Can't open $xlsfile. It may be in use by Excel.";
    }

    # Use binmode if "\n" is encoded as 2 bytes
    print $fh "\n";
    $fh->flush();
    if (-s $fh == 2) { binmode($fh) }
    seek($fh, 0, 0);

    # Store filehandle
    $self->{_filehandle} = $fh;

    # Check if "pack" gives the required IEEE 64bit float
    my $teststr = pack "d", 1.2345;
    my @hexdata = (0x8D, 0x97, 0x6E, 0x12, 0x83, 0xC0, 0xF3, 0x3F);
    my $number  = pack "C8", @hexdata;

    if ($number eq $teststr) {
        # Little Endian
        $self->{_byte_order} = 0;
    }
    elsif ($number eq reverse($teststr)){
        # Big Endian
        $self->{_byte_order} = 1;
    }
    else {
        # Give up. I'll fix this in a later version.
        croak ( "Required floating point format not supported "  .
                "on this platform. See the portability section " .
                "of the documentation."
        );
    }

    # Write binary header information
    $self->_xl_write_bof();
    $self->_xl_write_dimensions();
}


######################################################################
#
# Finalization routine
#

sub close {
    DESTROY;
}


######################################################################
#
# Destructor
#

sub DESTROY {
    my $self = shift;
    $self->_xl_write_eof();
    CORE::close($self->{_filehandle});
}


######################################################################
#
# _xl_write_bof()
#
# Writes Excel BOF record to indicate the beginning of a file
# in the compound document format.
#

sub _xl_write_bof {
    my $self      = shift;
    my $name      = 0x0809; # Record identifier
    my $length    = 0x0008; # Number of bytes to follow

    my $version   = 0x0005; # Excel BIFF version 5
    my $type      = 0x0010; # Worksheet
    my $build     = 0x0000; # Set to zero
    my $year      = 0x0000; # Set to zero

    my $header    = pack("vv",   $name, $length);
    my $data      = pack("vvvv", $version, $type, $build, $year);

    print {$self->{_filehandle}}  $header . $data;
}


######################################################################
#
# _xl_write_eof()
#
# Writes Excel EOF record to indicate the end of a file in the
# compound document format.
#

sub _xl_write_eof {
    my $self      = shift;
    my $name      = 0x000A; # Record identifier
    my $length    = 0x0000; # Number of bytes to follow

    my $header    = pack("vv", $name, $length);

    print {$self->{_filehandle}} $header;
}


######################################################################
#
# _xl_write_dimensions()
#
# Writes Excel DIMENSIONS to define the area in which there is data.
# Setting these values doesn't have an effect in this implementation.
#

sub _xl_write_dimensions {
    my $self      = shift;
    my $name      = 0x0000; # Record identifier
    my $length    = 0x000A; # Number of bytes to follow
    my $row_min   = 0;      # First row
    my $row_max   = 1;      # Last row plus 1
    my $col_min   = 0;      # First column
    my $col_max   = 1;      # Last column plus 1
    my $reserved  = 0x0000; # Reserved by Excel

    my $header    = pack("vv",    $name, $length);
    my $data      = pack("vvvvv", $row_min, $row_max,
                                  $col_min, $col_max, $reserved);

    print {$self->{_filehandle}} $header . $data;
}


######################################################################
#
# xl_write ($row, $col, $token)
#
# Parse $token as a number or string and call xl_write_number()
# or xl_write_string() accordingly. $row and $column are zero
# indexed.
#
# Returns: return value of called subroutine
#

sub xl_write {
    my $self  = shift;
    my $token = $_[2];

    # Match number or string
    if ($token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/){
        return $self->xl_write_number(@_);
    }
    else {
        return $self->xl_write_string(@_);
    }
}


######################################################################
#
# xl_write_number($row, $col, $num)
#
# Write a double to the specified row and column (zero indexed).
# An integer can be written as a double. Excel will display an
# integer.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#

sub xl_write_number {
    my $self      = shift;
    if (@_ < 3) { return -1 }
    
    my $name      = 0x0203; # Record identifier
    my $length    = 0x000E; # Number of bytes to follow

    my $row       = $_[0];  # Zero indexed row
    my $col       = $_[1];  # Zero indexed column

    if ($row >= $self->{_maxrow}) { return -2 }
    if ($col >= $self->{_maxcol}) { return -2 }

    my $xf        = 0x0000; # The cell format - not implemented here
    my $num       = $_[2];

    my $header    = pack("vv",  $name, $length);
    my $data      = pack("vvv", $row, $col, $xf);
    my $xl_double = pack("d",   $num);

    if ($self->{_byte_order}) { $xl_double = reverse $xl_double }

    print {$self->{_filehandle}} $header . $data . $xl_double;
    return 0;
}


######################################################################
#
# xl_write_string ($row, $col, $string)
#
# Write a string to the specified row and column (zero indexed).
# NOTE: there is an Excel 5 defined limit of 255 characters.
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long string truncated to 255 chars
#

sub xl_write_string {
    my $self      = shift;
    if (@_ < 3) { return -1 }

    my $name      = 0x0204; # Record identifier
    my $length    = 0x0008 + length($_[2]); # Num. of bytes to follow

    my $row       = $_[0];  # Zero indexed row
    my $col       = $_[1];  # Zero indexed column
    my $xf        = 0x0000; # The cell format - not implemented here
    my $strlen    = length($_[2]);
    my $str       = $_[2];
    my $str_error = 0;

    if ($row >= $self->{_maxrow}) { return -2 }
    if ($col >= $self->{_maxcol}) { return -2 }
    if ($strlen > $self->{_maxstr}) { # LABEL restricted to 255 chars
        $str       = substr($str, 0, $self->{_maxstr});
        $length    = 0x0008 + $self->{_maxstr};
        $strlen    = $self->{_maxstr};
        $str_error = -3;
    }

    my $header    = pack("vv",   $name, $length);
    my $data      = pack("vvvv", $row, $col, $xf, $strlen);

    print {$self->{_filehandle}} $header . $data . $str;
    return $str_error;
}

1;
__END__


=head1 NAME

Spreadsheet::WriteExcel - Write text and numbers to minimal Excel binary file.

=head1 VERSION

This document refers to version 0.08 of Spreadsheet::WriteExcel, released Jan 16, 2000.

=head1 SYNOPSIS

To write a string and a number in an Excel file called perl.xls:

    use Spreadsheet::WriteExcel;

    $row1 = $col1 = 0;
    $row2 = 1;

    $excel = Spreadsheet::WriteExcel->new("perl.xls");

    $excel->xl_write($row1, $col1, "Hi Excel!");
    $excel->xl_write($row2, $col1, 1.2345);

Or explicitly, without the overhead of parsing:

    $excel->xl_write_string($row1, $col1, "Hi Excel!");
    $excel->xl_write_number($row2, $col1, 1.2345);

The file is closed when the program ends or when it is no longer referred to. Alternatively you can close it as follows:

    $excel->close();

The following example converts a tab separated file called C<tab.txt> into an Excel file called C<tab.xls>.

    #!/usr/bin/perl -w

    use strict;
    use Spreadsheet::WriteExcel;

    open (TABFILE, "tab.txt") or die "tab.txt: $!";

    my $row = 0;
    my $col;

    my $excel = Spreadsheet::WriteExcel->new("tab.xls");

    while (<TABFILE>) {
        chomp;
        my @Fld = split('\t', $_);
        my $token;

        $col = 0;
        foreach $token (@Fld) {
           $excel->xl_write($row, $col, $token);
           $col++;
        }
        $row++;
    }


=head1 DESCRIPTION

=head2 Overview

This module can be used to write numbers and text in the native Excel binary file format. This is a minimal implementation of an Excel file; no formatting can be applied to cells and only a single worksheet can be written to a workbook.

It is intended to be cross-platform, however, this is not guaranteed. See the section on portability below.


=head2 Constructor and initialization

A new Excel file is created as follows:

    Spreadsheet::WriteExcel->new("filename.xls");

This will create a workbook called "filename.xls" with a single worksheet called "filename".

=head2 Object methods

The following are the methods provided by WriteExcel:

    xl_write(row, column, token)
    xl_write_number(row, column, number)
    xl_write_string(row, column, string)
    close()

Row and column are zero indexed cell locations; thus, Cell A1 is (0,0) and Cell AD2000 is (1999,29). Cells can be written to in any order. They can also be overwritten.

The method xl_write() calls xl_write_number() if "token" matches the following regex:

    $token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/

Otherwise it calls xl_write_string().

The write methods return:

    0 for success
   -1 for insufficient number of arguments
   -2 for row or column out of bounds
   -3 for string too long.
   

See also the section about limits.

The C<close()> method can be called to explicitly close the Excel file. Otherwise the file will be closed automatically when the object reference goes out of scope or the program ends.

=head2 Limits

The following limits are imposed by Excel or the version of the BIFF file that has been implemented:

    Description                          Limit   Source
    -----------------------------------  ------  -------
    Maximum number of chars in a string  255     Excel 5
    Maximum number of columns            256     Excel 5, 97
    Maximum number of rows in Excel 5    16,384  Excel 5
    Maximum number of rows in Excel 97   65,536  Excel 97

=head2 The Excel "BIFF" binary format

The binary format of an Excel file is referred to as the Excel "Binary Interchange File Format" (BIFF) file format. For details of this file format refer to the "Excel Developer's Kit", Microsoft Press. This module is based on the BIFF5 specification. To facilitate portability and ease of implementation the Compound Document wrapper is not used. This effectively limits the scope of the BIFF file to the records given below.

The following binary records are implemented:

    [BOF]
    [DIMENSIONS]
    [NUMBER]
    [LABEL]
    [EOF]

Each Excel BIFF binary record has the following format:

    Record name   - Identifier, 2 bytes
    Record length - Length of the subsequent data, 2 bytes
    Record data   - Data, variable length

=head1 PORTABILITY

WriteExcel.pm will only work on systems where perl packs floats in 64 bit IEEE format. The float must also be in little-endian format but WriteExcel.pm will reverse it as necessary.

Thus:

    print join(" ", map { sprintf "%#02x", $_ } unpack("C*", pack "d", 1.2345)), "\n";

should give (or in reverse order):

    0x8d 0x97 0x6e 0x12 0x83 0xc0 0xf3 0x3f

If your system doesn't support this format of float then WriteExcel will croak with the message given in the Diagnostics section. A future version will correct this, if possible. In the meantime, if this doesn't work for your OS let me know about it.

=head1 DIAGNOSTICS

=over 4

=item Filename required in WriteExcel("Filename")

A filename must be given in the constructor.

=item Can't open filename. It may be in use by Excel.

The file cannot be opened for writing or reading. It may be protected or already in use.

=item Required floating point format not supported on this platform.

Operating system doesn't support 64 bit IEEE float or it is byte-ordered in a way unknown to WriteExcel.

=back

=head1 ALTERNATIVES

Depending on your requirements, background and general sensibilities you may prefer one of the following methods of getting data into Excel:

* CSV, comma separated variables or text. If the file extension is C<csv>, Excel will open and convert this format automatically.

* HTML tables. This is an easy way of adding formatting.

* LAOLA. This is a Perl interface to OLE file formats, see CPAN.

* ODBC. Connect to an Excel file as a database.

* Office automation via the Win32::OLE module. This is very flexible and gives you access to multiple worksheets, formatting, and Excel's built-in functions.

=head1 BUGS

The main bug is the lack of a portable way of writing a little-endian 64 bit IEEE float. This is to-do.

=head1 AUTHOR

John McNamara (C<john.exeng@abanet.it>)

"Life, friends is boring. We must not say so." - John Berryman.

=head1 COPYRIGHT

Copyright (c) 1999-2000, John McNamara. All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
