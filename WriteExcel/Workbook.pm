package Spreadsheet::WriteExcel::Workbook;

###############################################################################
#
# Workbook - A writer class for Excel Workbooks.
#
#
# Used in conjunction with Spreadsheet::WriteExcel
#
# Copyright 2000-2001, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

require Exporter;

use strict;
use Carp;
use Spreadsheet::WriteExcel::BIFFwriter;
use Spreadsheet::WriteExcel::Worksheet;
use Spreadsheet::WriteExcel::OLEwriter;
use Spreadsheet::WriteExcel::Format;

use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::WriteExcel::BIFFwriter Exporter);

$VERSION = '0.08';

###############################################################################
#
# new()
#
# Constructor. Creates a new Workbook object from a BIFFwriter object.
#
sub new {

    my $class       = shift;
    my $filename    = $_[0] || '';
    my $self        = Spreadsheet::WriteExcel::BIFFwriter->new();
    my $ole_writer  = Spreadsheet::WriteExcel::OLEwriter->new($filename);
    my $tmp_format  = Spreadsheet::WriteExcel::Format->new();

    my $byte_order  = $self->{_byte_order};
    my $parser      = Spreadsheet::WriteExcel::Formula->new($byte_order);

    $self->{_OLEwriter}         = $ole_writer;
    $self->{_parser}            = $parser;
    $self->{_1904}              = 0;
    $self->{_activesheet}       = 0;
    $self->{_firstsheet}        = 0;
    $self->{_xf_index}          = 16; # 15 style XF's and 1 cell XF.
    $self->{_fileclosed}        = 0;
    $self->{_biffsize}          = 0;
    $self->{_sheetname}         = "Sheet";
    $self->{_tmp_format}        = $tmp_format;
    $self->{_url_format}        = '';
    $self->{_worksheets}        = [];
    $self->{_sheetnames}        = [];
    $self->{_formats}           = [];

    bless $self, $class;

    # Add the default format for hyperlinks
    my $url_format = $self->addformat();
    $url_format->set_color('blue');
    $url_format->set_underline(1);
    $self->{_url_format} = $url_format;

    $self->_tmpfile_warning();

    return $self;
}


###############################################################################
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


###############################################################################
#
# DESTROY()
#
# Close the workbook if it hasn't already been explicitly closed.
#
sub DESTROY {

    my $self = shift;

    $self->close() if not $self->{_fileclosed};
}


###############################################################################
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


###############################################################################
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
                        \$self->{_activesheet},
                        \$self->{_firstsheet},
                        $self->{_url_format},
                        $self->{_parser},
                    );

    my $worksheet = Spreadsheet::WriteExcel::Worksheet->new(@init_data);
    $self->{_worksheets}->[$index] = $worksheet;    # Store ref for iterator
    $self->{_sheetnames}->[$index] = $name;         # Store EXTERNSHEET names
    $self->{_parser}->set_ext_sheet($name, $index); # Store names in Formula.pm
    return $worksheet;
}


###############################################################################
#
# addformat()
#
# Add a new format to the Excel workbook. This adds an XF record and
# a FONT record.
#
sub addformat {

    my $self      = shift;

    my $format = Spreadsheet::WriteExcel::Format->new($self->{_xf_index});
    $self->{_xf_index} += 1;

    push @{$self->{_formats}}, $format;

    return $format;
}


###############################################################################
#
# set_1904()
#
# Set the date system: 0 = 1900 (the default), 1 = 1904
#
sub set_1904{

    my $self      = shift;

    if (defined($_[0])) {
        $self->{_1904} = $_[0];
    }
    else {
        $self->{_1904} = 1;
    }
}


###############################################################################
#
# get_1904()
#
# Return the date system: 0 = 1900, 1 = 1904
#
sub get_1904{

    my $self = shift;

    return $self->{_1904};
}


###############################################################################
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
    carp("Calling write() methods on a workbook object is deprecated," .
         " use write() in conjunction with a worksheet object instead"
        ) if $^W;
    return $self->{_worksheets}[0]->write(@_);
}


###############################################################################
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
    carp("Calling write() methods on a workbook object is deprecated," .
         " use write() in conjunction with a worksheet object instead"
        ) if $^W;
    return $self->{_worksheets}[0]->write_string(@_);
}


###############################################################################
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
    carp("Calling write() methods on a workbook object is deprecated," .
         " use write() in conjunction with a worksheet object instead"
        ) if $^W;
    return $self->{_worksheets}[0]->write_number(@_);
}


###############################################################################
#
# _tmpfile_warning()
#
# Check that tmp files can be created for use in Worksheet.pm. A CGI, mod_perl
# or IIS might not have permission to create tmp files. The test is here rather
# than in Worksheet.pm so that only one warning is given.
#
sub _tmpfile_warning{

    my $fh = IO::File->new_tmpfile();

    if ((not defined $fh) && ($^W)) {
        carp("Unable to create tmp files via IO::File->new_tmpfile(). " .
             "Storing data in memory ")
    }
}


###############################################################################
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
        $sheet->_close($self->{_sheetnames});
    }

    # Add Workbook globals
    $self->_store_bof(0x0005);
    $self->_store_window1();
    $self->_store_1904();
    $self->_store_all_fonts();
    $self->_store_all_num_formats();
    $self->_store_all_xfs();
    $self->_store_all_styles();
    $self->_calc_sheet_offsets();

    # Add BOUNDSHEET records
    foreach my $sheet (@{$self->{_worksheets}}) {
        $self->_store_boundsheet($sheet->{_name}, $sheet->{_offset});
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


###############################################################################
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
        $offset += $BOF + length($sheet->{_name});
    }

    $offset += $EOF;

    foreach my $sheet (@{$self->{_worksheets}}) {
        $sheet->{_offset} = $offset;
        $offset += $sheet->{_datasize};
    }

    $self->{_biffsize} = $offset;
}


###############################################################################
#
# _store_all_fonts()
#
# Store the Excel FONT records.
#
sub _store_all_fonts {

    my $self   = shift;

    # _tmp_format is added by new(). We use this to write the default XF's
    my $format = $self->{_tmp_format};
    my $font   = $format->get_font();

    # Note: Fonts are 0-indexed. According to the SDK there is no index 4,
    # so the following fonts are 0, 1, 2, 3, 5
    #
    for (1..5){
        $self->_append($font);
    }


    # Iterate through the XF objects and write a FONT record if it isn't the
    # same as the default FONT and if it hasn't already been used.
    #
    my %fonts;
    my $key;
    my $index = 6;                  # The first user defined FONT

    $key = $format->get_font_key(); # The default font from _tmp_format
    $fonts{$key} = 0;               # Index of the default font


    foreach $format (@{$self->{_formats}}) {
        $key = $format->get_font_key();

        if (exists $fonts{$key}) {
            # FONT has already been used
            $format->{_font_index} = $fonts{$key};
        }
        else {
            # Add a new FONT record
            $fonts{$key}           = $index;
            $format->{_font_index} = $index;
            $index++;
            $font = $format->get_font();
            $self->_append($font);
        }
    }
}


###############################################################################
#
# _store_all_num_formats()
#
# Store user defined numerical formats i.e. FORMAT records
#
sub _store_all_num_formats {

    my $self   = shift;

    # Leaning num_format syndrome
    my %num_formats;
    my @num_formats;
    my $num_format;
    my $index = 164;

    # Iterate through the XF objects and write a FORMAT record if it isn't a
    # built-in format type and if the FORMAT string hasn't already been used.
    #
    foreach my $format (@{$self->{_formats}}) {
        my $num_format = $format->{_num_format};

        # Check if $num_format is an index to a builtin format.
        # Also check for a string of zeros, which is a valid format string
        # but would evaluate to zero.
        #
        if ($num_format !~ m/^0+\d/) {
            next if $num_format =~ m/^\d+$/; # builtin
        }

        if (exists($num_formats{$num_format})) {
            # FORMAT has already been used
            $format->{_num_format} = $num_formats{$num_format};
        }
        else{
            # Add a new FORMAT
            $num_formats{$num_format} = $index;
            $format->{_num_format}    = $index;
            push @num_formats, $num_format;
            $index++;
        }
    }

    # Write the new FORMAT records starting from 0xA4
    $index = 164;
    foreach $num_format (@num_formats) {
        $self->_store_num_format($num_format, $index);
        $index++;
    }
}


###############################################################################
#
# _store_all_xfs()
#
# Write all XF records.
#
sub _store_all_xfs {

    my $self    = shift;

    # _tmp_format is added by new(). We use this to write the default XF's
    # The default font index is 0
    #
    my $format = $self->{_tmp_format};
    my $xf;

    for (0..14) {
        $xf = $format->get_xf(0xFFF5); # Style XF
        $self->_append($xf);
    }

    $xf = $format->get_xf(0x0001);     # Cell XF
    $self->_append($xf);


    # User defined XFs
    foreach $format (@{$self->{_formats}}) {
        $xf = $format->get_xf(0x0001);
        $self->_append($xf);
    }
}


###############################################################################
#
# _store_all_styles()
#
# Write all STYLE records.
#
sub _store_all_styles {

    my $self    = shift;

    $self->_store_style();
}


###############################################################################
#
# BIFF RECORDS
#


###############################################################################
#
# _store_window1()
#
# Write Excel BIFF WINDOW1 record.
#
sub _store_window1 {

    my $self      = shift;
    my $record    = 0x003D;  # Record identifier
    my $length    = 0x0012;  # Number of bytes to follow

    my $xWn       = 0x0000;  # Horizontal position of window
    my $yWn       = 0x0000;  # Vertical position of window
    my $dxWn      = 0x25BC;  # Width of window
    my $dyWn      = 0x1572;  # Height of window

    my $grbit     = 0x0038;  # Option flags
    my $ctabsel   = 0x0001;  # Number of workbook tabs selected
    my $wTabRatio = 0x0258;  # Tab to scrollbar ratio

    my $itabFirst = $self->{_firstsheet};  # 1st displayed worksheet
    my $itabCur   = $self->{_activesheet}; # Selected worksheet

    my $header    = pack("vv",        $record, $length);
    my $data      = pack("vvvvvvvvv", $xWn, $yWn, $dxWn, $dyWn,
                                      $grbit,
                                      $itabCur, $itabFirst,
                                      $ctabsel, $wTabRatio);

    $self->_append($header, $data);
}


###############################################################################
#
# _store_boundsheet()
#
# Writes Excel BIFF BOUNDSHEET record.
#
sub _store_boundsheet {

    my $self      = shift;

    my $record    = 0x0085;               # Record identifier
    my $length    = 0x07 + length($_[0]); # Number of bytes to follow

    my $sheetname = $_[0];                # Worksheet name
    my $offset    = $_[1];                # Location of worksheet BOF
    my $grbit     = 0x0000;               # Sheet identifier
    my $cch       = length($sheetname);   # Length of sheet name

    my $header    = pack("vv",  $record, $length);
    my $data      = pack("VvC", $offset, $grbit, $cch);

    $self->_append($header, $data, $sheetname);
}


###############################################################################
#
# _store_style()
#
# Write Excel BIFF STYLE records.
#
sub _store_style {

    my $self      = shift;

    my $record    = 0x0293; # Record identifier
    my $length    = 0x0004; # Bytes to follow

    my $ixfe      = 0x8000; # Index to style XF
    my $BuiltIn   = 0x00;   # Built-in style
    my $iLevel    = 0xff;   # Outline style level

    my $header    = pack("vv",  $record, $length);
    my $data      = pack("vCC", $ixfe, $BuiltIn, $iLevel);

    $self->_append($header, $data);
}


###############################################################################
#
# _store_num_format()
#
# Writes Excel FORMAT record for non "built-in" numerical formats.
#
sub _store_num_format {

    my $self      = shift;

    my $record    = 0x041E;                 # Record identifier
    my $length    = 0x03 + length($_[0]);   # Number of bytes to follow

    my $format    = $_[0];                  # Custom format string
    my $ifmt      = $_[1];                  # Format index code
    my $cch       = length($format);        # Length of format string

    my $header    = pack("vv", $record, $length);
    my $data      = pack("vC", $ifmt, $cch);

    $self->_append($header, $data, $format)
}


###############################################################################
#
# _store_1904()
#
# Write Excel 1904 record to indicate the date system in use.
#
sub _store_1904 {

    my $self      = shift;

    my $record    = 0x0022;         # Record identifier
    my $length    = 0x0002;         # Bytes to follow

    my $f1904     = $self->{_1904}; # Flag for 1904 date system

    my $header    = pack("vv",  $record, $length);
    my $data      = pack("v", $f1904);

    $self->_append($header, $data);
}


1;


__END__


=head1 NAME

Workbook - A writer class for Excel Workbooks.

=head1 SYNOPSIS

See the documentation for Spreadsheet::WriteExcel

=head1 DESCRIPTION

This module is used in conjunction with Spreadsheet::WriteExcel.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMI, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
