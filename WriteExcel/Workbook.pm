package Spreadsheet::WriteExcel::Workbook;

###############################################################################
#
# Workbook - A writer class for Excel Workbooks.
#
#
# Used in conjunction with Spreadsheet::WriteExcel
#
# Copyright 2000-2003, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

use Exporter;
use strict;
use Carp;
use Spreadsheet::WriteExcel::BIFFwriter;
use Spreadsheet::WriteExcel::OLEwriter;
use Spreadsheet::WriteExcel::Worksheet;
use Spreadsheet::WriteExcel::Format;


use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::WriteExcel::BIFFwriter Exporter);

$VERSION = '0.19';

###############################################################################
#
# new()
#
# Constructor. Creates a new Workbook object from a BIFFwriter object.
#
sub new {

    my $class       = shift;
    my $self        = Spreadsheet::WriteExcel::BIFFwriter->new();
    my $tmp_format  = Spreadsheet::WriteExcel::Format->new();
    my $byte_order  = $self->{_byte_order};
    my $parser      = Spreadsheet::WriteExcel::Formula->new($byte_order);

    $self->{_filename}          = $_[0] || '';
    $self->{_parser}            = $parser;
    $self->{_tempdir}           = undef;
    $self->{_1904}              = 0;
    $self->{_activesheet}       = 0;
    $self->{_firstsheet}        = 0;
    $self->{_selected}          = 0;
    $self->{_xf_index}          = 16; # 15 style XF's and 1 cell XF.
    $self->{_fileclosed}        = 0;
    $self->{_biffsize}          = 0;
    $self->{_sheetname}         = "Sheet";
    $self->{_tmp_format}        = $tmp_format;
    $self->{_url_format}        = '';
    $self->{_codepage}          = 0x04E4;
    $self->{_worksheets}        = [];
    $self->{_sheetnames}        = [];
    $self->{_formats}           = [];
    $self->{_palette}           = [];

    bless $self, $class;

    # Add the default format for hyperlinks
    $self->{_url_format} = $self->add_format(color => 'blue', underline => 1);


    # Check for a filename unless it is an existing filehandle
    if (not ref $self->{_filename} and $self->{_filename} eq '') {
        carp 'Filename required by Spreadsheet::WriteExcel->new()';
        return undef;
    }


    # Try to open the named file and see if it throws any errors.
    # If the filename is a reference it is assumed that it is a valid
    # filehandle and ignored
    #
    if (not ref $self->{_filename}) {
        my $fh = FileHandle->new('>'. $self->{_filename});
        if (not defined $fh) {
            carp "Can't open " .
                  $self->{_filename} .
                  ". It may be in use or protected";
            return undef;
        }
        $fh->close;
    }


    # Set colour palette.
    $self->set_palette_xl97();

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

    return if $self->{_fileclosed}; # Prevent close() from being called twice.

    $self->{_fileclosed} = 1;

    return $self->_store_workbook();
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
# sheets(slice,...)
#
# An accessor for the _worksheets[] array
#
# Returns: an optionally sliced list of the worksheet objects in a workbook.
#
sub sheets {

    my $self = shift;

    if (@_) {
        # Return a slice of the array
        return @{$self->{_worksheets}}[@_];
    }
    else {
        # Return the entire list
        return @{$self->{_worksheets}};
    }
}


###############################################################################
#
# worksheets()
#
# An accessor for the _worksheets[] array.
# This method is now deprecated. Use the sheets() method instead.
#
# Returns: an array reference
#
sub worksheets {

    my $self = shift;

    return $self->{_worksheets};
}


###############################################################################
#
# add_worksheet($name)
#
# Add a new worksheet to the Excel workbook.
# TODO: Add accessor for $self->{_sheetname} for international Excel versions.
#
# Returns: reference to a worksheet object
#
sub add_worksheet {

    my $self      = shift;
    my $name      = $_[0] || "";

    # Check that sheetname is <= 31 chars (Excel limit).
    croak "Sheetname $name must be <= 31 chars" if length $name > 31;

    # Check that sheetname doesn't contain any invalid characters
    croak 'Invalid Excel character [:*?/\\] in worksheet name: ' . $name
          if $name =~ m{[:*?/\\]};


    my $index     = @{$self->{_worksheets}};
    my $sheetname = $self->{_sheetname};

    if ($name eq "" ) { $name = $sheetname . ($index+1) }

    # Check that the worksheet name doesn't already exist: a fatal Excel error.
    foreach my $tmp (@{$self->{_worksheets}}) {
        croak "Worksheet '$name' already exists" if $name eq $tmp->get_name();
    }


    # Porters take note, the following scheme of passing references to Workbook
    # data (in the \$self->{_foo} cases) instead of a reference to the Workbook
    # itself is a workaround to avoid circular references between Workbook and
    # Worksheet objects. Feel free to implement this in any way the suits your
    # language.
    #
    my @init_data = (
                        $name,
                        $index,
                        \$self->{_activesheet},
                        \$self->{_firstsheet},
                        $self->{_url_format},
                        $self->{_parser},
                        $self->{_tempdir},
                    );

    my $worksheet = Spreadsheet::WriteExcel::Worksheet->new(@init_data);
    $self->{_worksheets}->[$index] = $worksheet;    # Store ref for iterator
    $self->{_sheetnames}->[$index] = $name;         # Store EXTERNSHEET names
    $self->{_parser}->set_ext_sheets($name, $index);# Store names in Formula.pm
    return $worksheet;
}


###############################################################################
#
# addworksheet($name)
#
# This method is now deprecated. Use the add_worksheet() method instead.
#
sub addworksheet {

    my $self = shift;

    $self->add_worksheet(@_);
}


###############################################################################
#
# add_format(%properties)
#
# Add a new format to the Excel workbook. This adds an XF record and
# a FONT record. Also, pass any properties to the Format::new().
#
sub add_format {

    my $self = shift;

    my $format = Spreadsheet::WriteExcel::Format->new($self->{_xf_index}, @_);

    $self->{_xf_index} += 1;
    push @{$self->{_formats}}, $format; # Store format reference

    return $format;
}


###############################################################################
#
# addformat()
#
# This method is now deprecated. Use the add_format() method instead.
#
sub addformat {

    my $self = shift;

    $self->add_format(@_);
}


###############################################################################
#
# set_1904()
#
# Set the date system: 0 = 1900 (the default), 1 = 1904
#
sub set_1904 {

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
sub get_1904 {

    my $self = shift;

    return $self->{_1904};
}


###############################################################################
#
# set_custom_color()
#
# Change the RGB components of the elements in the colour palette.
#
sub set_custom_color {

    my $self    = shift;


    # Match a HTML #xxyyzz style parameter
    if (defined $_[1] and $_[1] =~ /^#(\w\w)(\w\w)(\w\w)/ ) {
        @_ = ($_[0], hex $1, hex $2, hex $3);
    }


    my $index   = $_[0] || 0;
    my $red     = $_[1] || 0;
    my $green   = $_[2] || 0;
    my $blue    = $_[3] || 0;

    my $aref    = $self->{_palette};

    # Check that the colour index is the right range
    if ($index < 8 or $index > 64) {
        carp "Color index $index outside range: 8 <= index <= 64";
        return 0;
    }

    # Check that the colour components are in the right range
    if ( ($red   < 0 or $red   > 255) ||
         ($green < 0 or $green > 255) ||
         ($blue  < 0 or $blue  > 255) )
    {
        carp "Color component outside range: 0 <= color <= 255";
        return 0;
    }

    $index -=8; # Adjust colour index (wingless dragonfly)

    # Set the RGB value
    $$aref[$index] = [$red, $green, $blue, 0];

    return $index +8;
}


###############################################################################
#
# set_palette_xl97()
#
# Sets the colour palette to the Excel 97+ default.
#
sub set_palette_xl97 {

    my $self = shift;

    $self->{_palette} = [
                            [0x00, 0x00, 0x00, 0x00],   # 8
                            [0xff, 0xff, 0xff, 0x00],   # 9
                            [0xff, 0x00, 0x00, 0x00],   # 10
                            [0x00, 0xff, 0x00, 0x00],   # 11
                            [0x00, 0x00, 0xff, 0x00],   # 12
                            [0xff, 0xff, 0x00, 0x00],   # 13
                            [0xff, 0x00, 0xff, 0x00],   # 14
                            [0x00, 0xff, 0xff, 0x00],   # 15
                            [0x80, 0x00, 0x00, 0x00],   # 16
                            [0x00, 0x80, 0x00, 0x00],   # 17
                            [0x00, 0x00, 0x80, 0x00],   # 18
                            [0x80, 0x80, 0x00, 0x00],   # 19
                            [0x80, 0x00, 0x80, 0x00],   # 20
                            [0x00, 0x80, 0x80, 0x00],   # 21
                            [0xc0, 0xc0, 0xc0, 0x00],   # 22
                            [0x80, 0x80, 0x80, 0x00],   # 23
                            [0x99, 0x99, 0xff, 0x00],   # 24
                            [0x99, 0x33, 0x66, 0x00],   # 25
                            [0xff, 0xff, 0xcc, 0x00],   # 26
                            [0xcc, 0xff, 0xff, 0x00],   # 27
                            [0x66, 0x00, 0x66, 0x00],   # 28
                            [0xff, 0x80, 0x80, 0x00],   # 29
                            [0x00, 0x66, 0xcc, 0x00],   # 30
                            [0xcc, 0xcc, 0xff, 0x00],   # 31
                            [0x00, 0x00, 0x80, 0x00],   # 32
                            [0xff, 0x00, 0xff, 0x00],   # 33
                            [0xff, 0xff, 0x00, 0x00],   # 34
                            [0x00, 0xff, 0xff, 0x00],   # 35
                            [0x80, 0x00, 0x80, 0x00],   # 36
                            [0x80, 0x00, 0x00, 0x00],   # 37
                            [0x00, 0x80, 0x80, 0x00],   # 38
                            [0x00, 0x00, 0xff, 0x00],   # 39
                            [0x00, 0xcc, 0xff, 0x00],   # 40
                            [0xcc, 0xff, 0xff, 0x00],   # 41
                            [0xcc, 0xff, 0xcc, 0x00],   # 42
                            [0xff, 0xff, 0x99, 0x00],   # 43
                            [0x99, 0xcc, 0xff, 0x00],   # 44
                            [0xff, 0x99, 0xcc, 0x00],   # 45
                            [0xcc, 0x99, 0xff, 0x00],   # 46
                            [0xff, 0xcc, 0x99, 0x00],   # 47
                            [0x33, 0x66, 0xff, 0x00],   # 48
                            [0x33, 0xcc, 0xcc, 0x00],   # 49
                            [0x99, 0xcc, 0x00, 0x00],   # 50
                            [0xff, 0xcc, 0x00, 0x00],   # 51
                            [0xff, 0x99, 0x00, 0x00],   # 52
                            [0xff, 0x66, 0x00, 0x00],   # 53
                            [0x66, 0x66, 0x99, 0x00],   # 54
                            [0x96, 0x96, 0x96, 0x00],   # 55
                            [0x00, 0x33, 0x66, 0x00],   # 56
                            [0x33, 0x99, 0x66, 0x00],   # 57
                            [0x00, 0x33, 0x00, 0x00],   # 58
                            [0x33, 0x33, 0x00, 0x00],   # 59
                            [0x99, 0x33, 0x00, 0x00],   # 60
                            [0x99, 0x33, 0x66, 0x00],   # 61
                            [0x33, 0x33, 0x99, 0x00],   # 62
                            [0x33, 0x33, 0x33, 0x00],   # 63
                        ];

    return 0;
}


###############################################################################
#
# set_palette_xl5()
#
# Sets the colour palette to the Excel 5 default.
#
sub set_palette_xl5 {

    my $self = shift;

    $self->{_palette} = [
                            [0x00, 0x00, 0x00, 0x00],   # 8
                            [0xff, 0xff, 0xff, 0x00],   # 9
                            [0xff, 0x00, 0x00, 0x00],   # 10
                            [0x00, 0xff, 0x00, 0x00],   # 11
                            [0x00, 0x00, 0xff, 0x00],   # 12
                            [0xff, 0xff, 0x00, 0x00],   # 13
                            [0xff, 0x00, 0xff, 0x00],   # 14
                            [0x00, 0xff, 0xff, 0x00],   # 15
                            [0x80, 0x00, 0x00, 0x00],   # 16
                            [0x00, 0x80, 0x00, 0x00],   # 17
                            [0x00, 0x00, 0x80, 0x00],   # 18
                            [0x80, 0x80, 0x00, 0x00],   # 19
                            [0x80, 0x00, 0x80, 0x00],   # 20
                            [0x00, 0x80, 0x80, 0x00],   # 21
                            [0xc0, 0xc0, 0xc0, 0x00],   # 22
                            [0x80, 0x80, 0x80, 0x00],   # 23
                            [0x80, 0x80, 0xff, 0x00],   # 24
                            [0x80, 0x20, 0x60, 0x00],   # 25
                            [0xff, 0xff, 0xc0, 0x00],   # 26
                            [0xa0, 0xe0, 0xe0, 0x00],   # 27
                            [0x60, 0x00, 0x80, 0x00],   # 28
                            [0xff, 0x80, 0x80, 0x00],   # 29
                            [0x00, 0x80, 0xc0, 0x00],   # 30
                            [0xc0, 0xc0, 0xff, 0x00],   # 31
                            [0x00, 0x00, 0x80, 0x00],   # 32
                            [0xff, 0x00, 0xff, 0x00],   # 33
                            [0xff, 0xff, 0x00, 0x00],   # 34
                            [0x00, 0xff, 0xff, 0x00],   # 35
                            [0x80, 0x00, 0x80, 0x00],   # 36
                            [0x80, 0x00, 0x00, 0x00],   # 37
                            [0x00, 0x80, 0x80, 0x00],   # 38
                            [0x00, 0x00, 0xff, 0x00],   # 39
                            [0x00, 0xcf, 0xff, 0x00],   # 40
                            [0x69, 0xff, 0xff, 0x00],   # 41
                            [0xe0, 0xff, 0xe0, 0x00],   # 42
                            [0xff, 0xff, 0x80, 0x00],   # 43
                            [0xa6, 0xca, 0xf0, 0x00],   # 44
                            [0xdd, 0x9c, 0xb3, 0x00],   # 45
                            [0xb3, 0x8f, 0xee, 0x00],   # 46
                            [0xe3, 0xe3, 0xe3, 0x00],   # 47
                            [0x2a, 0x6f, 0xf9, 0x00],   # 48
                            [0x3f, 0xb8, 0xcd, 0x00],   # 49
                            [0x48, 0x84, 0x36, 0x00],   # 50
                            [0x95, 0x8c, 0x41, 0x00],   # 51
                            [0x8e, 0x5e, 0x42, 0x00],   # 52
                            [0xa0, 0x62, 0x7a, 0x00],   # 53
                            [0x62, 0x4f, 0xac, 0x00],   # 54
                            [0x96, 0x96, 0x96, 0x00],   # 55
                            [0x1d, 0x2f, 0xbe, 0x00],   # 56
                            [0x28, 0x66, 0x76, 0x00],   # 57
                            [0x00, 0x45, 0x00, 0x00],   # 58
                            [0x45, 0x3e, 0x01, 0x00],   # 59
                            [0x6a, 0x28, 0x13, 0x00],   # 60
                            [0x85, 0x39, 0x6a, 0x00],   # 61
                            [0x4a, 0x32, 0x85, 0x00],   # 62
                            [0x42, 0x42, 0x42, 0x00],   # 63
                        ];

    return 0;
}


###############################################################################
#
# set_tempdir()
#
# Change the default temp directory used by _initialize() in Worksheet.pm.
#
sub set_tempdir {

    my $self = shift;

    # Windows workaround. See Worksheet::_initialize()
    my $dir  = shift || '';

    croak "$dir is not a valid directory"       if $dir ne '' and not -d $dir;
    croak "set_tempdir must be called before add_worksheet" if $self->sheets();

    $self->{_tempdir} = $dir ;
}


###############################################################################
#
# set_codepage()
#
# See also the _store_codepage method. This is used to store the code page, i.e.
# the character set used in the workbook.
#
sub set_codepage {

    my $self        = shift;
    my $codepage    = $_[0] || 1;
    $codepage   = 0x04E4 if $codepage == 1;
    $codepage   = 0x8000 if $codepage == 2;
    $self->{_codepage} = $codepage;
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

    # Ensure that at least one worksheet has been selected.
    if ($self->{_activesheet} == 0) {
        @{$self->{_worksheets}}[0]->{_selected} = 1;
    }

    # Calculate the number of selected worksheet tabs and call the finalization
    # methods for each worksheet
    foreach my $sheet (@{$self->{_worksheets}}) {
        $self->{_selected}++ if $sheet->{_selected};
        $sheet->_close($self->{_sheetnames});
    }

    # Add Workbook globals
    $self->_store_bof(0x0005);
    $self->_store_codepage();
    $self->_store_externs();    # For print area and repeat rows
    $self->_store_names();      # For print area and repeat rows
    $self->_store_window1();
    $self->_store_1904();
    $self->_store_all_fonts();
    $self->_store_all_num_formats();
    $self->_store_all_xfs();
    $self->_store_all_styles();
    $self->_store_palette();
    $self->_calc_sheet_offsets();

    # Add BOUNDSHEET records
    foreach my $sheet (@{$self->{_worksheets}}) {
        $self->_store_boundsheet($sheet->{_name}, $sheet->{_offset});
    }

    # End Workbook globals
    $self->_store_eof();

    # Store the workbook in an OLE container
    return $self->_store_OLE_file();
}


###############################################################################
#
# _store_OLE_file()
#
# Store the workbook in an OLE container if the total size of the workbook data
# is less than ~ 7MB.
#
sub _store_OLE_file {

    my $self = shift;

    my $OLE  = Spreadsheet::WriteExcel::OLEwriter->new($self->{_filename});

    # Write Worksheet data if data <~ 7MB
    if ($OLE->set_size($self->{_biffsize})) {
        $OLE->write_header();
        $OLE->write($self->{_data});

        foreach my $sheet (@{$self->{_worksheets}}) {
            while (my $tmp = $sheet->get_data()) {
                $OLE->write($tmp);
            }
        }

        return $OLE->close();
    }
    else {
        # File in greater than limit, set $! to "File too large"
        $! = 27; # Perl error code "File too large"
        my $maxsize = 7_087_104;

        croak "Maximum Spreadsheet::WriteExcel filesize, $maxsize bytes, "    .
              "exceeded. To create files bigger than this limit please refer ".
              "to the \"Spreadsheet::WriteExcel::Big\" documentation.\n"      ;

        # return 0;
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

    my $self = shift;

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

    my $self = shift;

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

        # Check if $num_format is an index to a built-in format.
        # Also check for a string of zeros, which is a valid format string
        # but would evaluate to zero.
        #
        if ($num_format !~ m/^0+\d/) {
            next if $num_format =~ m/^\d+$/; # built-in
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

    my $self = shift;

    # _tmp_format is added by new(). We use this to write the default XF's
    # The default font index is 0
    #
    my $format = $self->{_tmp_format};
    my $xf;

    for (0..14) {
        $xf = $format->get_xf('style'); # Style XF
        $self->_append($xf);
    }

    $xf = $format->get_xf('cell');      # Cell XF
    $self->_append($xf);


    # User defined XFs
    foreach $format (@{$self->{_formats}}) {
        $xf = $format->get_xf('cell');
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

    my $self = shift;

    $self->_store_style();
}


###############################################################################
#
# _store_externs()
#
# Write the EXTERNCOUNT and EXTERNSHEET records. These are used as indexes for
# the NAME records.
#
sub _store_externs {

    my $self = shift;

    # Create EXTERNCOUNT with number of worksheets
    $self->_store_externcount(scalar @{$self->{_worksheets}});

    # Create EXTERNSHEET for each worksheet
    foreach my $sheetname (@{$self->{_sheetnames}}) {
        $self->_store_externsheet($sheetname);
    }
}


###############################################################################
#
# _store_names()
#
# Write the NAME record to define the print area and the repeat rows and cols.
#
sub _store_names {

    my $self = shift;

    # Create the print area NAME records
    foreach my $worksheet (@{$self->{_worksheets}}) {
        # Write a Name record if the print area has been defined
        if (defined $worksheet->{_print_rowmin}) {
            $self->_store_name_short(
                $worksheet->{_index},
                0x06, # NAME type
                $worksheet->{_print_rowmin},
                $worksheet->{_print_rowmax},
                $worksheet->{_print_colmin},
                $worksheet->{_print_colmax}
            );
        }
    }


    # Create the print title NAME records
    foreach my $worksheet (@{$self->{_worksheets}}) {

        my $rowmin = $worksheet->{_title_rowmin};
        my $rowmax = $worksheet->{_title_rowmax};
        my $colmin = $worksheet->{_title_colmin};
        my $colmax = $worksheet->{_title_colmax};

        # Determine if row + col, row, col or nothing has been defined
        # and write the appropriate record
        #
        if (defined $rowmin && defined $colmin) {
            # Row and column titles have been defined.
            # Row title has been defined.
            $self->_store_name_long(
                $worksheet->{_index},
                0x07, # NAME type
                $rowmin,
                $rowmax,
                $colmin,
                $colmax
           );
        }
        elsif (defined $rowmin) {
            # Row title has been defined.
            $self->_store_name_short(
                $worksheet->{_index},
                0x07, # NAME type
                $rowmin,
                $rowmax,
                0x00,
                0xff
            );
        }
        elsif (defined $colmin) {
            # Column title has been defined.
            $self->_store_name_short(
                $worksheet->{_index},
                0x07, # NAME type
                0x0000,
                0x3fff,
                $colmin,
                $colmax
            );
        }
        else {
            # Print title hasn't been defined.
        }
    }
}




###############################################################################
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

    my $record    = 0x003D;                 # Record identifier
    my $length    = 0x0012;                 # Number of bytes to follow

    my $xWn       = 0x0000;                 # Horizontal position of window
    my $yWn       = 0x0000;                 # Vertical position of window
    my $dxWn      = 0x25BC;                 # Width of window
    my $dyWn      = 0x1572;                 # Height of window

    my $grbit     = 0x0038;                 # Option flags
    my $ctabsel   = $self->{_selected};     # Number of workbook tabs selected
    my $wTabRatio = 0x0258;                 # Tab to scrollbar ratio

    my $itabFirst = $self->{_firstsheet};   # 1st displayed worksheet
    my $itabCur   = $self->{_activesheet};  # Active worksheet

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

    $self->_append($header, $data, $format);
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


###############################################################################
#
# _store_externcount($count)
#
# Write BIFF record EXTERNCOUNT to indicate the number of external sheet
# references in the workbook.
#
# Excel only stores references to external sheets that are used in NAME.
# The workbook NAME record is required to define the print area and the repeat
# rows and columns.
#
# A similar method is used in Worksheet.pm for a slightly different purpose.
#
sub _store_externcount {

    my $self     = shift;

    my $record   = 0x0016;          # Record identifier
    my $length   = 0x0002;          # Number of bytes to follow

    my $cxals    = $_[0];           # Number of external references

    my $header   = pack("vv", $record, $length);
    my $data     = pack("v",  $cxals);

    $self->_append($header, $data);
}


###############################################################################
#
# _store_externsheet($sheetname)
#
#
# Writes the Excel BIFF EXTERNSHEET record. These references are used by
# formulas. NAME record is required to define the print area and the repeat
# rows and columns.
#
# A similar method is used in Worksheet.pm for a slightly different purpose.
#
sub _store_externsheet {

    my $self        = shift;

    my $record      = 0x0017;               # Record identifier
    my $length      = 0x02 + length($_[0]); # Number of bytes to follow

    my $sheetname   = $_[0];                # Worksheet name
    my $cch         = length($sheetname);   # Length of sheet name
    my $rgch        = 0x03;                 # Filename encoding

    my $header      = pack("vv",  $record, $length);
    my $data        = pack("CC", $cch, $rgch);

    $self->_append($header, $data, $sheetname);
}


###############################################################################
#
# _store_name_short()
#
#
# Store the NAME record in the short format that is used for storing the print
# area, repeat rows only and repeat columns only.
#
sub _store_name_short {

    my $self            = shift;

    my $record          = 0x0018;       # Record identifier
    my $length          = 0x0024;       # Number of bytes to follow

    my $index           = shift;        # Sheet index
    my $type            = shift;

    my $grbit           = 0x0020;       # Option flags
    my $chKey           = 0x00;         # Keyboard shortcut
    my $cch             = 0x01;         # Length of text name
    my $cce             = 0x0015;       # Length of text definition
    my $ixals           = $index +1;    # Sheet index
    my $itab            = $ixals;       # Equal to ixals
    my $cchCustMenu     = 0x00;         # Length of cust menu text
    my $cchDescription  = 0x00;         # Length of description text
    my $cchHelptopic    = 0x00;         # Length of help topic text
    my $cchStatustext   = 0x00;         # Length of status bar text
    my $rgch            = $type;        # Built-in name type

    my $unknown03       = 0x3b;
    my $unknown04       = 0xffff-$index;
    my $unknown05       = 0x0000;
    my $unknown06       = 0x0000;
    my $unknown07       = 0x1087;
    my $unknown08       = 0x8005;

    my $rowmin          = $_[0];        # Start row
    my $rowmax          = $_[1];        # End row
    my $colmin          = $_[2];        # Start column
    my $colmax          = $_[3];        # end column


    my $header          = pack("vv",  $record, $length);
    my $data            = pack("v", $grbit);
    $data              .= pack("C", $chKey);
    $data              .= pack("C", $cch);
    $data              .= pack("v", $cce);
    $data              .= pack("v", $ixals);
    $data              .= pack("v", $itab);
    $data              .= pack("C", $cchCustMenu);
    $data              .= pack("C", $cchDescription);
    $data              .= pack("C", $cchHelptopic);
    $data              .= pack("C", $cchStatustext);
    $data              .= pack("C", $rgch);
    $data              .= pack("C", $unknown03);
    $data              .= pack("v", $unknown04);
    $data              .= pack("v", $unknown05);
    $data              .= pack("v", $unknown06);
    $data              .= pack("v", $unknown07);
    $data              .= pack("v", $unknown08);
    $data              .= pack("v", $index);
    $data              .= pack("v", $index);
    $data              .= pack("v", $rowmin);
    $data              .= pack("v", $rowmax);
    $data              .= pack("C", $colmin);
    $data              .= pack("C", $colmax);

    $self->_append($header, $data);
}


###############################################################################
#
# _store_name_long()
#
#
# Store the NAME record in the long format that is used for storing the repeat
# rows and columns when both are specified. This share a lot of code with
# _store_name_short() but we use a separate method to keep the code clean.
# Code abstraction for reuse can be carried too far, and I should know. ;-)
#
sub _store_name_long {

    my $self            = shift;

    my $record          = 0x0018;       # Record identifier
    my $length          = 0x003d;       # Number of bytes to follow

    my $index           = shift;        # Sheet index
    my $type            = shift;

    my $grbit           = 0x0020;       # Option flags
    my $chKey           = 0x00;         # Keyboard shortcut
    my $cch             = 0x01;         # Length of text name
    my $cce             = 0x002e;       # Length of text definition
    my $ixals           = $index +1;    # Sheet index
    my $itab            = $ixals;       # Equal to ixals
    my $cchCustMenu     = 0x00;         # Length of cust menu text
    my $cchDescription  = 0x00;         # Length of description text
    my $cchHelptopic    = 0x00;         # Length of help topic text
    my $cchStatustext   = 0x00;         # Length of status bar text
    my $rgch            = $type;        # Built-in name type

    my $unknown01       = 0x29;
    my $unknown02       = 0x002b;
    my $unknown03       = 0x3b;
    my $unknown04       = 0xffff-$index;
    my $unknown05       = 0x0000;
    my $unknown06       = 0x0000;
    my $unknown07       = 0x1087;
    my $unknown08       = 0x8008;

    my $rowmin          = $_[0];        # Start row
    my $rowmax          = $_[1];        # End row
    my $colmin          = $_[2];        # Start column
    my $colmax          = $_[3];        # end column


    my $header          = pack("vv",  $record, $length);
    my $data            = pack("v", $grbit);
    $data              .= pack("C", $chKey);
    $data              .= pack("C", $cch);
    $data              .= pack("v", $cce);
    $data              .= pack("v", $ixals);
    $data              .= pack("v", $itab);
    $data              .= pack("C", $cchCustMenu);
    $data              .= pack("C", $cchDescription);
    $data              .= pack("C", $cchHelptopic);
    $data              .= pack("C", $cchStatustext);
    $data              .= pack("C", $rgch);
    $data              .= pack("C", $unknown01);
    $data              .= pack("v", $unknown02);
    # Column definition
    $data              .= pack("C", $unknown03);
    $data              .= pack("v", $unknown04);
    $data              .= pack("v", $unknown05);
    $data              .= pack("v", $unknown06);
    $data              .= pack("v", $unknown07);
    $data              .= pack("v", $unknown08);
    $data              .= pack("v", $index);
    $data              .= pack("v", $index);
    $data              .= pack("v", 0x0000);
    $data              .= pack("v", 0x3fff);
    $data              .= pack("C", $colmin);
    $data              .= pack("C", $colmax);
    # Row definition
    $data              .= pack("C", $unknown03);
    $data              .= pack("v", $unknown04);
    $data              .= pack("v", $unknown05);
    $data              .= pack("v", $unknown06);
    $data              .= pack("v", $unknown07);
    $data              .= pack("v", $unknown08);
    $data              .= pack("v", $index);
    $data              .= pack("v", $index);
    $data              .= pack("v", $rowmin);
    $data              .= pack("v", $rowmax);
    $data              .= pack("C", 0x00);
    $data              .= pack("C", 0xff);
    # End of data
    $data              .= pack("C", 0x10);

    $self->_append($header, $data);
}


###############################################################################
#
# _store_palette()
#
# Stores the PALETTE biff record.
#
sub _store_palette {

    my $self            = shift;

    my $aref            = $self->{_palette};

    my $record          = 0x0092;           # Record identifier
    my $length          = 2 + 4 * @$aref;   # Number of bytes to follow
    my $ccv             =         @$aref;   # Number of RGB values to follow
    my $data;                               # The RGB data

    # Pack the RGB data
    $data .= pack "CCCC", @$_ for @$aref;

    my $header = pack("vvv",  $record, $length, $ccv);

    $self->_append($header, $data);
}


###############################################################################
#
# _store_codepage()
#
# Stores the CODEPAGE biff record.
#
sub _store_codepage {

    my $self            = shift;

    my $record          = 0x0042;               # Record identifier
    my $length          = 0x0002;               # Number of bytes to follow
    my $cv              = $self->{_codepage};   # The code page

    my $header          = pack("vv", $record, $length);
    my $data            = pack("v",  $cv);

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

© MM-MMIII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
