package Spreadsheet::WriteExcel;

######################################################################
#
# WriteExcel - Write text and numbers to minimal Excel binary file.
#
# Copyright 1999-2000, John McNamara, writeexcel@eircom.net
#
# Documentation after __END__
#

require Exporter;

use strict;
use Carp;
use FileHandle;

use vars qw($VERSION @ISA);

@ISA = qw(Exporter);
$VERSION = '0.10.00'; # 13 May 2000, cummings


######################################################################
#
# Constructor
#
sub new {
    my $class  = $_[0];
    my $rowmax = 65536;
    my $colmax = 256;
    my $strmax = 255;

    my $self   = {
                    _xlsfilename   => $_[1] || "",
                    _filehandle    => "",
                    _fileclosed     => 0,
                    _byte_order    => "",
                    _xls_rowmax    => $rowmax,
                    _xls_colmax    => $colmax,
                    _xls_strmax    => $strmax,
                    _dim_rowmin    => $rowmax +1,
                    _dim_rowmax    => 0,
                    _dim_colmin    => $colmax +1,
                    _dim_colmax    => 0,
                    _dim_offset    => 0,
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

    # Open file for writing
    my $fh = FileHandle->new("> $xlsfile");
    if (not defined $fh) {
        croak "Can't open $xlsfile. It may be in use by Excel.";
    }

    # binmode file whether platform requires it or not
    binmode($fh);

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
    $self->_write_bof();
    $self->_write_dimensions();
}


######################################################################
#
# Finalization routine
#
sub close {
    my $self = shift;
    $self->_write_eof();
    CORE::close($self->{_filehandle});
    $self->{_fileclosed} = 1;
}


######################################################################
#
# Destructor
#
sub DESTROY {
    my $self = shift;
    if (not $self->{_fileclosed}) { $self->close() }
}


######################################################################
#
# _write_bof()
#
# Writes Excel BOF record to indicate the beginning of a file
# in the compound document format.
#
sub _write_bof {
    my $self      = shift;
    my $name      = 0x0809; # Record identifier
    my $length    = 0x0008; # Number of bytes to follow

    # Use Biff2 version to avoid "Previous version" warnings in Excel
    my $version   = 0x0000; # Should be 0x0500 for Biff 5
    my $type      = 0x0010; # Worksheet
    my $build     = 0x0000; # Set to zero
    my $year      = 0x0000; # Set to zero

    my $header    = pack("vv",   $name, $length);
    my $data      = pack("vvvv", $version, $type, $build, $year);

    print {$self->{_filehandle}}  $header . $data;
}


######################################################################
#
# _write_dimensions()
#
# Writes Excel DIMENSIONS to define the area in which there is data.
# The initial default values are overwritten before closing the file.
#
sub _write_dimensions {
    my $self      = shift;
    my $name      = 0x0000;          # Record identifier
    my $length    = 0x000A;          # Number of bytes to follow
    my $row_min   = ($_[0] || 0);    # First row
    my $row_max   = ($_[1] || 0) +1; # Last row plus 1
    my $col_min   = ($_[2] || 0);    # First column
    my $col_max   = ($_[3] || 0) +1; # Last column plus 1
    my $reserved  = 0x0000;          # Reserved by Excel

    my $header    = pack("vv",    $name, $length);
    my $data      = pack("vvvvv", $row_min, $row_max,
                                  $col_min, $col_max, $reserved);

    # Store offset of DIMENSIONS record
    $self->{_dim_offset} = tell($self->{_filehandle});

    print {$self->{_filehandle}} $header . $data;
}


######################################################################
#
# _write_eof()
#
# Writes Excel EOF record to indicate the end of a file in the
# compound document format. Rewrite DIMENSIONS record.
#
sub _write_eof {
    my $self      = shift;
    my $row_min   = $self->{_dim_rowmin};
    my $row_max   = $self->{_dim_rowmax};
    my $col_min   = $self->{_dim_colmin};
    my $col_max   = $self->{_dim_colmax};

    if ($row_min == $self->{_xls_rowmax} +1) { $row_min = 0 };
    if ($col_min == $self->{_xls_colmax} +1) { $col_min = 0 };

    my $name      = 0x000A; # Record identifier
    my $length    = 0x0000; # Number of bytes to follow

    my $header    = pack("vv", $name, $length);

    print {$self->{_filehandle}} $header;

    # Rewrite DIMENSIONS record with correct range
    seek($self->{_filehandle}, $self->{_dim_offset}, 0);
    $self->_write_dimensions($row_min, $row_max, $col_min, $col_max);
}


######################################################################
#
# write ($row, $col, $token)
#
# Parse $token as a number or string and call write_number()
# or write_string() accordingly. $row and $column are zero
# indexed.
#
# Returns: return value of called subroutine
#
sub write {
    my $self  = shift;
    my $token = $_[2];

    # Match number or string
    if ($token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/){
        return $self->write_number(@_);
    }
    else {
        return $self->write_string(@_);
    }
}


######################################################################
#
# write_number($row, $col, $num)
#
# Write a double to the specified row and column (zero indexed).
# An integer can be written as a double. Excel will display an
# integer.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub write_number {
    my $self      = shift;
    if (@_ < 3) { return -1 }

    my $name      = 0x0203; # Record identifier
    my $length    = 0x000E; # Number of bytes to follow

    my $row       = $_[0];  # Zero indexed row
    my $col       = $_[1];  # Zero indexed column
    my $xf        = 0x0000; # The cell format - not implemented here
    my $num       = $_[2];

    if ($row >= $self->{_xls_rowmax}) { return -2 }
    if ($col >= $self->{_xls_colmax}) { return -2 }
    if ($row <  $self->{_dim_rowmin}) { $self->{_dim_rowmin} = $row }
    if ($row >  $self->{_dim_rowmax}) { $self->{_dim_rowmax} = $row }
    if ($col <  $self->{_dim_colmin}) { $self->{_dim_colmin} = $col }
    if ($col >  $self->{_dim_colmax}) { $self->{_dim_colmax} = $col }

    my $header    = pack("vv",  $name, $length);
    my $data      = pack("vvv", $row, $col, $xf);
    my $xl_double = pack("d",   $num);

    if ($self->{_byte_order}) { $xl_double = reverse $xl_double }

    print {$self->{_filehandle}} $header . $data . $xl_double;
    return 0;
}


######################################################################
#
# write_string ($row, $col, $string)
#
# Write a string to the specified row and column (zero indexed).
# NOTE: there is an Excel 5 defined limit of 255 characters.
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long string truncated to 255 chars
#
sub write_string {
    my $self      = shift;
    if (@_ < 3) { return -1 }

    my $name      = 0x0204; # Record identifier
    my $length    = 0x0008 + length($_[2]); # Bytes to follow

    my $row       = $_[0];  # Zero indexed row
    my $col       = $_[1];  # Zero indexed column
    my $xf        = 0x0000; # The cell format - not implemented here
    my $strlen    = length($_[2]);
    my $str       = $_[2];
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

    my $header    = pack("vv",   $name, $length);
    my $data      = pack("vvvv", $row, $col, $xf, $strlen);

    print {$self->{_filehandle}} $header . $data . $str;
    return $str_error;
}


#####################################################################
#
# Routines to map older version method names to new method names
#
sub xl_write {
    my $self  = shift;
    return $self->write(@_);
}

sub xl_write_string {
    my $self  = shift;
    return $self->write_string(@_);
}

sub xl_write_number {
    my $self  = shift;
    return $self->write_number(@_);
}


1;

__END__


=head1 NAME

Spreadsheet::WriteExcel - Write text and numbers to minimal Excel binary file.

=head1 VERSION

This document refers to version 0.10.00 of Spreadsheet::WriteExcel, released May 13, 2000.

=head1 SYNOPSIS

To write a string and a number to an Excel file called perl.xls:

    use Spreadsheet::WriteExcel;

    $row1 = $col1 = 0;
    $row2 = 1;

    $excel = Spreadsheet::WriteExcel->new("perl.xls");

    $excel->write($row1, $col1, "Hi Excel!");
    $excel->write($row2, $col1, 1.2345);

Or explicitly, without the overhead of parsing:

    $excel->write_string($row1, $col1, "Hi Excel!");
    $excel->write_number($row2, $col1, 1.2345);

The file is closed when the program ends or when it is no longer referred to. Alternatively you can close it as follows:

    $excel->close();

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

    write(row, column, token)
    write_number(row, column, number)
    write_string(row, column, string)
    close()

Row and column are zero indexed cell locations; thus, Cell A1 is (0,0) and Cell AD2000 is (1999,29). Cells can be written to in any order. They can also be overwritten. (QuickView users refer to the bugs section.)

The method write() calls write_number() if "token" matches the following regex:

    $token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/

Otherwise it calls write_string().

The C<write> methods return:

    0 for success
   -1 for insufficient number of arguments
   -2 for row or column out of bounds
   -3 for string too long.


See also the section about limits.

The C<close()> method can be called to explicitly close the Excel file. Otherwise the file will be closed automatically when the object reference goes out of scope or the program ends.

Note: The write* methods were previously named xl_write*. The older method names are still available but deprecated.

=head2 Example

The following example converts a tab separated file called C<tab.txt> into an Excel file called C<tab.xls>.

    #!/usr/bin/perl -w

    use strict;
    use Spreadsheet::WriteExcel;

    open (TABFILE, "tab.txt") or die "tab.txt: $!";

    my $excel = Spreadsheet::WriteExcel->new("tab.xls");
    my $row = 0;
    my $col;

    while (<TABFILE>) {
        chomp;
        my @Fld = split('\t', $_);

        $col = 0;
        foreach my $token (@Fld) {
           $excel->write($row, $col, $token);
           $col++;
        }
        $row++;
    }

=head2 Limits

The following limits are imposed by Excel or the version of the BIFF file that has been implemented:

    Description                          Limit   Source
    -----------------------------------  ------  -------
    Maximum number of chars in a string  255     Excel 5
    Maximum number of columns            256     Excel 5, 97
    Maximum number of rows in Excel 5    16,384  Excel 5
    Maximum number of rows in Excel 97   65,536  Excel 97

=head2 The Excel BIFF binary format

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

If your system doesn't support this format of float then WriteExcel will croak with the message given in the Diagnostics section. A future version will correct this, if possible. 

=head1 DIAGNOSTICS

=over 4

=item Filename required in WriteExcel('Filename')

A filename must be given in the constructor.

=item Can't open filename. It may be in use by Excel.

The file cannot be opened for writing. It may be protected or already in use.

=item Required floating point format not supported on this platform.

Operating system doesn't support 64 bit IEEE float or it is byte-ordered in a way unknown to WriteExcel.

=back

=head1 WRITING EXCEL FILES

Depending on your requirements, background and general sensibilities you may prefer one of the following methods of getting data into Excel:

* CSV, comma separated variables or text. If the file extension is C<csv>, Excel will open and convert this format automatically.

* HTML tables. This is an easy way of adding formatting.

* ODBC. Connect to an Excel file as a database.

* Win32::OLE module and office automation. This requires a Windows platform and an installed copy of Excel. However, it is easy to use and gives access to the complete range of Excel's features such as: multiple worksheets, charts, cell formatting, macros and the built-in functions. See http://www.activestate.com/ActivePerl/docs/faq/Windows/ActivePerl-Winfaq12.html and http://www.activestate.com/ActivePerl/docs/site/lib/Win32/OLE.html

=head1 READING EXCEL FILES

Despite the title of this module the most commonly asked questions are in relation to reading Excel files. To read data from Excel files try:

* HTML tables. If the files are saved from Excel in a HTML format the data can be accessed using HTML::TableExtract http://search.cpan.org/search?dist=HTML-TableExtract

* ODBC.

* OLE::Storage, aka LAOLA. This is a Perl interface to OLE file formats. In particular, the distro contains an Excel to HTML converter called Herbert, http://user.cs.tu-berlin.de/~schwartz/pmh/ There is also an open source C/C++ project based on the LAOLA work. Try the Filters Project at http://arturo.directmail.org/filtersweb/ and the xlHtml Project at http://www.gate.net/~ddata/xlHtml/index.htm The xlHtml filter is more complete than Herbert.

* Win32::OLE module and office automation. This requires a Windows platform and an installed copy of Excel. This is the most powerful and complete method for interfacing with Excel. See http://www.activestate.com/ActivePerl/docs/faq/Windows/ActivePerl-Winfaq12.html and http://www.activestate.com/ActivePerl/docs/site/lib/Win32/OLE.html

Also, if you wish to view Excel files on Windows platforms which don't have Excel installed you can use the free Microsoft Excel Viewer http://officeupdate.microsoft.com/downloadDetails/xlviewer.htm

=head1 BUGS

The main bug is the lack of a portable way of writing a little-endian 64 bit IEEE float. This is to-do.

Other Spreadsheets: The binary file created by WriteExcel is not a complete Excel file. As a result it is not compatible with XESS, Applix, Star Office or anything else. This may be fixed indirectly in a later version. In the meantime write to non-Excel spreadsheets in their native format.

QuickView: Excel files written with Version 0.08 are not displayed correctly in MS or JASC QuickView. This is partially fixed in Version 0.09 onwards. However, if you wish to write files that are fully compatible with QuickView it is necessary to write the cells in a sequential row by row order. This does not apply to Excel or to Excel Viewer.

=head1 TO DO

If possible, this module will be extended to include multiple worksheets, and formatting for rows, columns and cells.

=head1 ACKNOWLEDGEMENTS

The following people contributed to the debugging and testing of WriteExcel.pm:

Arthur@ais, Mike Blazer, CPAN testers, Johan Ekenberg, Paul J. Falbe, Artur Silveira da Cunha, John Wren.


=head1 AUTHOR

John McNamara writeexcel@eircom.net

    Buffalo Bill's
    defunct
            who used to
            ride a watersmooth-silver
                                            stallion
    and break onetwothreefourfive pigeonsjustlikethat
                                                                Jesus
    he was a handsome man
                                    and what i want to know is
    how do you like your blueeyed boy
    Mister Death
    --e.e. cummings

=head1 COPYRIGHT

Copyright (c) 2000, John McNamara. All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
