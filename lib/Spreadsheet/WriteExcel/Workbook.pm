package Spreadsheet::WriteExcel::Workbook;

###############################################################################
#
# Workbook - A writer class for Excel Workbooks.
#
#
# Used in conjunction with Spreadsheet::WriteExcel
#
# Copyright 2000-2006, John McNamara, jmcnamara@cpan.org
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
use Spreadsheet::WriteExcel::Chart;

use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::WriteExcel::BIFFwriter Exporter);

$VERSION = '2.16';

###############################################################################
#
# new()
#
# Constructor. Creates a new Workbook object from a BIFFwriter object.
#
sub new {

    my $class       = shift;
    my $self        = Spreadsheet::WriteExcel::BIFFwriter->new();
    my $byte_order  = $self->{_byte_order};
    my $parser      = Spreadsheet::WriteExcel::Formula->new($byte_order);

    $self->{_filename}          = $_[0] || '';
    $self->{_parser}            = $parser;
    $self->{_tempdir}           = undef;
    $self->{_1904}              = 0;
    $self->{_activesheet}       = 0;
    $self->{_firstsheet}        = 0;
    $self->{_selected}          = 0;
    $self->{_xf_index}          = 0;
    $self->{_fileclosed}        = 0;
    $self->{_biffsize}          = 0;
    $self->{_sheetname}         = "Sheet";
    $self->{_url_format}        = '';
    $self->{_codepage}          = 0x04E4;
    $self->{_worksheets}        = [];
    $self->{_sheetnames}        = [];
    $self->{_formats}           = [];
    $self->{_palette}           = [];

    $self->{_using_tmpfile}     = 1;
    $self->{_filehandle}        = "";
    $self->{_temp_file}         = "";
    $self->{_internal_fh}       = 0;
    $self->{_fh_out}            = "";

    $self->{_str_total}         = 0;
    $self->{_str_unique}        = 0;
    $self->{_str_table}         = {};
    $self->{_str_array}         = [];
    $self->{_str_block_sizes}   = [];

    $self->{_ext_ref_count}     = 0;
    $self->{_ext_refs}          = {};

    $self->{_mso_clusters}      = [];
    $self->{_mso_size}          = 0;

    $self->{_hideobj}           = 0;


    bless $self, $class;


    # Add the in-built style formats and the default cell format.
    $self->add_format(type => 1);                       #  0 Normal
    $self->add_format(type => 1);                       #  1 RowLevel 1
    $self->add_format(type => 1);                       #  2 RowLevel 2
    $self->add_format(type => 1);                       #  3 RowLevel 3
    $self->add_format(type => 1);                       #  4 RowLevel 4
    $self->add_format(type => 1);                       #  5 RowLevel 5
    $self->add_format(type => 1);                       #  6 RowLevel 6
    $self->add_format(type => 1);                       #  7 RowLevel 7
    $self->add_format(type => 1);                       #  8 ColLevel 1
    $self->add_format(type => 1);                       #  9 ColLevel 2
    $self->add_format(type => 1);                       # 10 ColLevel 3
    $self->add_format(type => 1);                       # 11 ColLevel 4
    $self->add_format(type => 1);                       # 12 ColLevel 5
    $self->add_format(type => 1);                       # 13 ColLevel 6
    $self->add_format(type => 1);                       # 14 ColLevel 7
    $self->add_format();                                # 15 Cell XF
    $self->add_format(type => 1, num_format => 0x2B);   # 16 Comma
    $self->add_format(type => 1, num_format => 0x29);   # 17 Comma[0]
    $self->add_format(type => 1, num_format => 0x2C);   # 18 Currency
    $self->add_format(type => 1, num_format => 0x2A);   # 19 Currency[0]
    $self->add_format(type => 1, num_format => 0x09);   # 20 Percent


    # Add the default format for hyperlinks
    $self->{_url_format} = $self->add_format(color => 'blue', underline => 1);


    # Check for a filename unless it is an existing filehandle
    if (not ref $self->{_filename} and $self->{_filename} eq '') {
        carp 'Filename required by Spreadsheet::WriteExcel->new()';
        return undef;
    }


    # Convert the filename to a filehandle to pass to the OLE writer when the
    # file is closed. If the filename is a reference it is assumed that it is
    # a valid filehandle.
    #
    if (not ref $self->{_filename}) {

        my $fh = FileHandle->new('>'. $self->{_filename});

        if (not defined $fh) {
            carp "Can't open " .
                  $self->{_filename} .
                  ". It may be in use or protected";
            return undef;
        }

        # binmode file whether platform requires it or not
        binmode($fh);
        $self->{_internal_fh} = 1;
        $self->{_fh_out}      = $fh;
    }
    else {
        $self->{_internal_fh} = 0;
        $self->{_fh_out}      = $self->{_filename};

    }


    # Set colour palette.
    $self->set_palette_xl97();

    $self->_initialize();
    return $self;
}


###############################################################################
#
# _initialize()
#
# Open a tmp file to store the majority of the Worksheet data. If this fails,
# for example due to write permissions, store the data in memory. This can be
# slow for large files.
#
# TODO: Move this and other methods shared with Worksheet up into BIFFWriter.
#
sub _initialize {

    my $self = shift;
    my $fh;
    my $tmp_dir;

    # The following code is complicated by Windows limitations. Porters can
    # choose a more direct method.



    # In the default case we use IO::File->new_tmpfile(). This may fail, in
    # particular with IIS on Windows, so we allow the user to specify a temp
    # directory via File::Temp.
    #
    if (defined $self->{_tempdir}) {

        # Delay loading File:Temp to reduce the module dependencies.
        eval { require File::Temp };
        die "The File::Temp module must be installed in order ".
            "to call set_tempdir().\n" if $@;


        # Trap but ignore File::Temp errors.
        eval { $fh = File::Temp::tempfile(DIR => $self->{_tempdir}) };

        # Store the failed tmp dir in case of errors.
        $tmp_dir = $self->{_tempdir} || File::Spec->tmpdir if not $fh;
    }
    else {

        $fh = IO::File->new_tmpfile();

        # Store the failed tmp dir in case of errors.
        $tmp_dir = "POSIX::tmpnam() directory" if not $fh;
    }


    # Check if the temp file creation was successful. Else store data in memory.
    if ($fh) {

        # binmode file whether platform requires it or not.
        binmode($fh);

        # Store filehandle
        $self->{_filehandle} = $fh;
    }
    else {

        # Set flag to store data in memory if XX::tempfile() failed.
        $self->{_using_tmpfile} = 0;

        if ($^W) {
            my $dir = $self->{_tempdir} || File::Spec->tmpdir();

            warn "Unable to create temp files in $tmp_dir. Data will be ".
                 "stored in memory. Refer to set_tempdir() in the ".
                 "Spreadsheet::WriteExcel documentation.\n" ;
        }
    }
}


###############################################################################
#
# _append(), overloaded.
#
# Store Worksheet data in memory using the base class _append() or to a
# temporary file, the default.
#
sub _append {

    my $self = shift;
    my $data = '';

    if ($self->{_using_tmpfile}) {
        $data = join('', @_);

        # Add CONTINUE records if necessary
        $data = $self->_add_continue($data) if length($data) > $self->{_limit};

        # Protect print() from -l on the command line.
        local $\ = undef;

        print {$self->{_filehandle}} $data;
        $self->{_datasize} += length($data);
    }
    else {
        $data = $self->SUPER::_append(@_);
    }

    return $data;
}


###############################################################################
#
# get_data().
#
# Retrieves data from memory in one chunk, or from disk in $buffer
# sized chunks.
#
sub get_data {

    my $self   = shift;
    my $buffer = 4096;
    my $tmp;

    # Return data stored in memory
    if (defined $self->{_data}) {
        $tmp           = $self->{_data};
        $self->{_data} = undef;
        my $fh         = $self->{_filehandle};
        seek($fh, 0, 0) if $self->{_using_tmpfile};
        return $tmp;
    }

    # Return data stored on disk
    if ($self->{_using_tmpfile}) {
        return $tmp if read($self->{_filehandle}, $tmp, $buffer);
    }

    # No data to return
    return undef;
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
# add_worksheet($name, $encoding)
#
# Add a new worksheet to the Excel workbook.
#
# Returns: reference to a worksheet object
#
sub add_worksheet {

    my $self     = shift;
    my $index    = @{$self->{_worksheets}};

    my ($name, $encoding) = $self->_check_sheetname($_[0], $_[1]);


    # Porters take note, the following scheme of passing references to Workbook
    # data (in the \$self->{_foo} cases) instead of a reference to the Workbook
    # itself is a workaround to avoid circular references between Workbook and
    # Worksheet objects. Feel free to implement this in any way the suits your
    # language.
    #
    my @init_data = (
                         $name,
                         $index,
                         $encoding,
                        \$self->{_activesheet},
                        \$self->{_firstsheet},
                         $self->{_url_format},
                         $self->{_parser},
                         $self->{_tempdir},
                        \$self->{_str_total},
                        \$self->{_str_unique},
                        \$self->{_str_table},
                         $self->{_1904},
                    );

    my $worksheet = Spreadsheet::WriteExcel::Worksheet->new(@init_data);
    $self->{_worksheets}->[$index] = $worksheet;     # Store ref for iterator
    $self->{_sheetnames}->[$index] = $name;          # Store EXTERNSHEET names
    $self->{_parser}->set_ext_sheets($name, $index); # Store names in Formula.pm
    return $worksheet;
}


###############################################################################
#
# add_chart_ext($name, $filename)
#
# Add an externally created chart.
#
#
sub add_chart_ext {

    my $self     = shift;
    my $filename = $_[0];
    my $index    = @{$self->{_worksheets}};

    my ($name, $encoding) = $self->_check_sheetname($_[1], $_[2]);


    my @init_data = (
                         $filename,
                         $name,
                         $index,
                         $encoding,
                        \$self->{_activesheet},
                        \$self->{_firstsheet},
                    );

    my $worksheet = Spreadsheet::WriteExcel::Chart->new(@init_data);
    $self->{_worksheets}->[$index] = $worksheet;     # Store ref for iterator
    $self->{_sheetnames}->[$index] = $name;          # Store EXTERNSHEET names
    $self->{_parser}->set_ext_sheets($name, $index); # Store names in Formula.pm
    return $worksheet;
}


###############################################################################
#
# _check_sheetname($name, $encoding)
#
# Check for valid worksheet names. We check the length, if it contains any
# invalid characters and if the name is unique in the workbook.
#
sub _check_sheetname {

    my $self            = shift;
    my $name            = $_[0] || "";
    my $encoding        = $_[1] || 0;
    my $limit           = $encoding ? 62 : 31;
    my $invalid_char    = qr([\[\]:*?/\\]);

    # Supply default "Sheet" name if none has been defined.
    my $index     = @{$self->{_worksheets}};
    my $sheetname = $self->{_sheetname};

    if ($name eq "" ) {
        $name     = $sheetname . ($index+1);
        $encoding = 0;
    }


    # Check that sheetname is <= 31 (1 or 2 byte chars). Excel limit.
    croak "Sheetname $name must be <= 31 chars" if length $name > $limit;

    # Check that Unicode sheetname has an even number of bytes
    croak 'Odd number of bytes in Unicode worksheet name:' . $name
          if $encoding == 1 and length($name) % 2;


    # Check that sheetname doesn't contain any invalid characters
    if ($encoding != 1 and $name =~ $invalid_char) {
        # Check ASCII names
        croak 'Invalid character []:*?/\\ in worksheet name: ' . $name;
    }
    else {
        # Extract any 8bit clean chars from the UTF16 name and validate them.
        for my $wchar ($name =~ /../sg) {
            my ($hi, $lo) = unpack "aa", $wchar;
            if ($hi eq "\0" and $lo =~ $invalid_char) {
                croak 'Invalid character []:*?/\\ in worksheet name: ' . $name;
            }
        }
    }


    # Handle utf8 strings in perl 5.8.
    if ($] >= 5.008) {
        require Encode;

        if (Encode::is_utf8($name)) {
            $name = Encode::encode("UTF-16BE", $name);
            $encoding = 1;
        }
    }


    # Check that the worksheet name doesn't already exist since this is a fatal
    # error in Excel 97. The check must also exclude case insensitive matches
    # since the names 'Sheet1' and 'sheet1' are equivalent. The tests also have
    # to take the encoding into account.
    #
    foreach my $worksheet (@{$self->{_worksheets}}) {
        my $name_a  = $name;
        my $encd_a  = $encoding;
        my $name_b  = $worksheet->{_name};
        my $encd_b  = $worksheet->{_encoding};
        my $error   = 0;

        if    ($encd_a == 0 and $encd_b == 0) {
            $error  = 1 if lc($name_a) eq lc($name_b);
        }
        elsif ($encd_a == 0 and $encd_b == 1) {
            $name_a = pack "n*", unpack "C*", $name_a;
            $error  = 1 if lc($name_a) eq lc($name_b);
        }
        elsif ($encd_a == 1 and $encd_b == 0) {
            $name_b = pack "n*", unpack "C*", $name_b;
            $error  = 1 if lc($name_a) eq lc($name_b);
        }
        elsif ($encd_a == 1 and $encd_b == 1) {
            # We can do a true case insensitive test with Perl 5.8 and utf8.
            if ($] >= 5.008) {
                $name_a = Encode::decode("UTF-16BE", $name_a);
                $name_b = Encode::decode("UTF-16BE", $name_b);
                $error  = 1 if lc($name_a) eq lc($name_b);
            }
            else {
            # We can't easily do a case insensite test of the UTF16 names.
            # As a special case we check if all of the high bytes are nulls and
            # then do an ASCII style case insensitive test.

                # Strip out the high bytes (funkily).
                my $hi_a = grep {ord} $name_a =~ /(.)./sg;
                my $hi_b = grep {ord} $name_b =~ /(.)./sg;

                if ($hi_a or $hi_b) {
                    $error  = 1 if    $name_a  eq    $name_b;
                }
                else {
                    $error  = 1 if lc($name_a) eq lc($name_b);
                }
            }
        }

        # If any of the cases failed we throw the error here.
        if ($error) {
            croak "Worksheet name '$name', with case ignored, " .
                  "is already in use";
        }
    }

    return ($name,  $encoding);
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

    croak "set_1904() must be called before add_worksheet" if $self->sheets();


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

    # Add a default worksheet if non have been added.
    $self->add_worksheet() if not @{$self->{_worksheets}};

    # Calculate size required for MSO records and update worksheets.
    $self->_calc_mso_sizes();

    # Ensure that at least one worksheet has been selected.
    if ($self->{_activesheet} == 0) {
        @{$self->{_worksheets}}[0]->{_selected} = 1;
        @{$self->{_worksheets}}[0]->{_hidden}   = 0;
    }

    # Calculate the number of selected worksheet tabs and call the finalization
    # methods for each worksheet
    foreach my $sheet (@{$self->{_worksheets}}) {
        $self->{_selected}++ if $sheet->{_selected};
        $sheet->{_active} = 1 if $sheet->{_index} == $self->{_activesheet};
        $sheet->_close($self->{_sheetnames});
    }

    # Add Workbook globals
    $self->_store_bof(0x0005);
    $self->_store_codepage();
    $self->_store_hideobj();
    $self->_store_window1();
    $self->_store_1904();
    $self->_store_all_fonts();
    $self->_store_all_num_formats();
    $self->_store_all_xfs();
    $self->_store_all_styles();
    $self->_store_palette();

    # Calculate the offsets required by the BOUNDSHEET records
    $self->_calc_sheet_offsets();

    # Add BOUNDSHEET records. For BIFF 7+ TODO ....
    foreach my $sheet (@{$self->{_worksheets}}) {
        $self->_store_boundsheet($sheet->{_name},
                                 $sheet->{_offset},
                                 $sheet->{_type},
                                 $sheet->{_hidden},
                                 $sheet->{_encoding});
    }

    # NOTE: If any records are added between here and EOF the
    # _calc_sheet_offsets() should be updated to include the new length.
    if ($self->{_ext_ref_count}) {
        $self->_store_supbook();
        $self->_store_externsheet();
        $self->_store_names();
    }
    $self->_add_mso_drawing_group();
    $self->_store_shared_strings();

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

    my $OLE  = Spreadsheet::WriteExcel::OLEwriter->new($self->{_fh_out});

    # Indicate that we created the filehandle and want to close it.
    $OLE->{_internal_fh} = $self->{_internal_fh};

    # Write Worksheet data if data <~ 7MB
    if ($OLE->set_size($self->{_biffsize})) {
        $OLE->write_header();

        while (my $tmp = $self->get_data()) {
            $OLE->write($tmp);
        }

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
# Calculate Worksheet BOF offsets records for use in the BOUNDSHEET records.
#
sub _calc_sheet_offsets {

    my $self    = shift;
    my $BOF     = 12;
    my $EOF     = 4;
    my $offset  = $self->{_datasize};

    # Add the length of the SST and associated CONTINUEs
    $offset += $self->_calculate_shared_string_sizes();

    # Add the length of the SUPBOOK, EXTERNSHEET and NAME records
    $offset += $self->_calculate_extern_sizes();

    # Add the length of the MSODRAWINGGROUP records
    $offset += $self->{_mso_size};


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
# _calc_mso_sizes()
#
# Calculate the MSODRAWINGGROUP sizes and the indexes of the Worksheet
# MSODRAWING records.
#
# In the following SPID is shape id, according to Escher nonemclature.
#
sub _calc_mso_sizes {

    my $self            = shift;

    my $num_shapes      = 0;
    my $mso_size        = 0;    # Size of the MSODRAWINGGROUP record
    my $start_spid      = 1024; # Initial spid for each sheet
    my $max_spid        = 1024; # spidMax
    my $num_clusters    = 1;    # cidcl
    my $shapes_saved    = 0;    # cspSaved
    my $drawings_saved  = 0;    # cdgSaved
    my @clusters        = ();


    # Iterate through the worksheets, calculate the MSODRAWINGGROUP parameters
    # and space required to store the record and the MSODRAWING parameters
    # required by each worksheet.
    #
    foreach my $sheet (@{$self->{_worksheets}}) {
        next unless $num_shapes = $sheet->_prepare_comments();

        # Include 1 parent MSODRAWING shape, per sheet, in the shape count.
        $num_shapes   += 1;
        $shapes_saved += $num_shapes;


        # Add a drawing object for each sheet with comments.
        $drawings_saved++;


        # For each sheet start the spids at the next 1024 interval.
        $max_spid   = 1024 * (1 + int(($max_spid -1)/1024));
        $start_spid = $max_spid;


        # Max spid for each sheet and eventually for the workbook.
        $max_spid  += $num_shapes;


        # Store the cluster ids
        for (my $i = $num_shapes; $i > 0; $i -= 1024) {
            $num_clusters  += 1;
            $mso_size      += 8;
            my $size        = $i > 1024 ? 1024 : $i;

            push @clusters, [$drawings_saved, $size];
        }


        # Pass calculate values back to the worksheet
        $sheet->{_comments_ids} = [$start_spid, $drawings_saved,
                                  $num_shapes, $max_spid -1];
    }


    # Calculate the MSODRAWINGGROUP size if we have stored some shapes.
    $mso_size              += 86 if $mso_size; # Smallest size is 86+8=94


    $self->{_mso_size}      = $mso_size;
    $self->{_mso_clusters}  = [
                                $max_spid, $num_clusters, $shapes_saved,
                                $drawings_saved, [@clusters]
                              ];
}


###############################################################################
#
# _store_all_fonts()
#
# Store the Excel FONT records.
#
sub _store_all_fonts {

    my $self    = shift;

    my $format  = $self->{_formats}->[15]; # The default cell format.
    my $font    = $format->get_font();

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

    $key = $format->get_font_key(); # The default font for cell formats.
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

    my %num_formats;
    my @num_formats;
    my $num_format;
    my $index = 164; # User defined FORMAT records start from 0xA4


    # Iterate through the XF objects and write a FORMAT record if it isn't a
    # built-in format type and if the FORMAT string hasn't already been used.
    #
    foreach my $format (@{$self->{_formats}}) {
        my $num_format = $format->{_num_format};
        my $encoding   = $format->{_num_format_enc};

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
            $self->_store_num_format($num_format, $index, $encoding);
            $index++;
        }
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

    foreach my $format (@{$self->{_formats}}) {
        my $xf = $format->get_xf();
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

    # Excel adds the built-in styles in alphabetical order.
    my @built_ins = (
        [0x03, 16], # Comma
        [0x06, 17], # Comma[0]
        [0x04, 18], # Currency
        [0x07, 19], # Currency[0]
        [0x00,  0], # Normal
        [0x05, 20], # Percent

        # We don't deal with these styles yet.
        #[0x08, 21], # Hyperlink
        #[0x02,  8], # ColLevel_n
        #[0x01,  1], # RowLevel_n
    );


    for my $aref (@built_ins) {
        my $type     = $aref->[0];
        my $xf_index = $aref->[1];

        $self->_store_style($type, $xf_index);
    }
}


###############################################################################
#
# _store_names()
#
# Write the NAME record to define the print area and the repeat rows and cols.
#
sub _store_names {

    my $self        = shift;
    my $index       = 0;
    my %ext_refs    = %{$self->{_ext_refs}};

    # Create the print area NAME records
    foreach my $worksheet (@{$self->{_worksheets}}) {

        my $key = "$index:$index";
        my $ref = $ext_refs{$key};
        $index++;

        # Write a Name record if the print area has been defined
        if (defined $worksheet->{_print_rowmin}) {
            $self->_store_name_short(
                $worksheet->{_index},
                0x06, # NAME type
                $ref,
                $worksheet->{_print_rowmin},
                $worksheet->{_print_rowmax},
                $worksheet->{_print_colmin},
                $worksheet->{_print_colmax}
            );
        }
    }

    $index = 0;

    # Create the print title NAME records
    foreach my $worksheet (@{$self->{_worksheets}}) {

        my $rowmin = $worksheet->{_title_rowmin};
        my $rowmax = $worksheet->{_title_rowmax};
        my $colmin = $worksheet->{_title_colmin};
        my $colmax = $worksheet->{_title_colmax};
        my $key    = "$index:$index";
        my $ref    = $ext_refs{$key};
        $index++;

        # Determine if row + col, row, col or nothing has been defined
        # and write the appropriate record
        #
        if (defined $rowmin && defined $colmin) {
            # Row and column titles have been defined.
            # Row title has been defined.
            $self->_store_name_long(
                $worksheet->{_index},
                0x07, # NAME type
                $ref,
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
                $ref,
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
                $ref,
                0x0000,
                0xffff,
                $colmin,
                $colmax
            );
        }
        else {
            # Nothing left to do
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
    my $length    = 0x08 + length($_[0]); # Number of bytes to follow

    my $sheetname = $_[0];                # Worksheet name
    my $offset    = $_[1];                # Location of worksheet BOF
    my $type      = $_[2];                # Worksheet type
    my $hidden    = $_[3];                # Worksheet hidden flag
    my $encoding  = $_[4];                # Sheet name encoding
    my $cch       = length($sheetname);   # Length of sheet name

    my $grbit     = $type | $hidden;

    # Character length is num of chars not num of bytes
    $cch /= 2 if $encoding;

    # Change the UTF-16 name from BE to LE
    $sheetname = pack 'n*', unpack 'v*', $sheetname if $encoding;

    my $header    = pack("vv",   $record, $length);
    my $data      = pack("VvCC", $offset, $grbit, $cch, $encoding);

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

    my $type      = $_[0];  # Built-in style
    my $xf_index  = $_[1];  # Index to style XF
    my $level     = 0xff;   # Outline style level

    $xf_index    |= 0x8000; # Add flag to indicate built-in style.


    my $header    = pack("vv",  $record, $length);
    my $data      = pack("vCC", $xf_index, $type, $level);

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

    my $record    = 0x041E;         # Record identifier
    my $length;                     # Number of bytes to follow

    my $format    = $_[0];          # Custom format string
    my $ifmt      = $_[1];          # Format index code
    my $encoding  = $_[2];          # Char encoding for format string


    # Handle utf8 strings in perl 5.8.
    if ($] >= 5.008) {
        require Encode;

        if (Encode::is_utf8($format)) {
            $format = Encode::encode("UTF-16BE", $format);
            $encoding = 1;
        }
    }


    # Char length of format string
    my $cch = length $format;


    # Handle Unicode format strings.
    if ($encoding == 1) {
        croak "Uneven number of bytes in Unicode font name" if $cch % 2;
        $cch    /= 2 if $encoding;
        $format  = pack 'v*', unpack 'n*', $format;
    }


    # Special case to handle Euro symbol, 0x80, in non-Unicode strings.
    if ($encoding == 0 and $format =~ /\x80/) {
        $format   =  pack 'v*', unpack 'C*', $format;
        $format   =~ s/\x80\x00/\xAC\x20/g;
        $encoding =  1;
    }

    $length       = 0x05 + length $format;

    my $header    = pack("vv", $record, $length);
    my $data      = pack("vvC", $ifmt, $cch, $encoding);

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
# _store_supbook()
#
# Write BIFF record SUPBOOK to indicate that the workbook contains external
# references, in our case, formula, print area and print title refs.
#
sub _store_supbook {

    my $self        = shift;

    my $record      = 0x01AE;                   # Record identifier
    my $length      = 0x0004;                   # Number of bytes to follow

    my $ctabs       = @{$self->{_worksheets}};  # Number of worksheets
    my $StVirtPath  = 0x0401;                   # Encoded workbook filename

    my $header      = pack("vv", $record, $length);
    my $data        = pack("vv", $ctabs, $StVirtPath);

    $self->_append($header, $data);
}


###############################################################################
#
# _store_externsheet()
#
#
# Writes the Excel BIFF EXTERNSHEET record. These references are used by
# formulas. TODO NAME record is required to define the print area and the repeat
# rows and columns.
#
sub _store_externsheet {

    my $self        = shift;

    my $record      = 0x0017;                   # Record identifier
    my $length;                                 # Number of bytes to follow


    # Get the external refs
    my %ext_refs = %{$self->{_ext_refs}};
    my @ext_refs = sort {$ext_refs{$a} <=> $ext_refs{$b}} keys %ext_refs;

    # Change the external refs from stringified "1:1" to [1, 1]
    foreach my $ref (@ext_refs) {
        $ref = [split /:/, $ref];
    }


    my $cxti        = scalar @ext_refs;         # Number of Excel XTI structures
    my $rgxti       = '';                       # Array of XTI structures

    # Write the XTI structs
    foreach my $ext_ref (@ext_refs) {
        $rgxti .= pack("vvv", 0, $ext_ref->[0], $ext_ref->[1])
    }


    my $data        = pack("v", $cxti) . $rgxti;
    my $header      = pack("vv", $record, length $data);

    $self->_append($header, $data);
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
    my $length          = 0x001b;       # Number of bytes to follow

    my $index           = shift;        # Sheet index
    my $type            = shift;
    my $ext_ref         = shift;        # TODO

    my $grbit           = 0x0020;       # Option flags
    my $chKey           = 0x00;         # Keyboard shortcut
    my $cch             = 0x01;         # Length of text name
    my $cce             = 0x000b;       # Length of text definition
    my $unknown01       = 0x0000;       #
    my $ixals           = $index +1;    # Sheet index
    my $unknown02       = 0x00;         #
    my $cchCustMenu     = 0x00;         # Length of cust menu text
    my $cchDescription  = 0x00;         # Length of description text
    my $cchHelptopic    = 0x00;         # Length of help topic text
    my $cchStatustext   = 0x00;         # Length of status bar text
    my $rgch            = $type;        # Built-in name type
    my $unknown03       = 0x3b;         #

    my $rowmin          = $_[0];        # Start row
    my $rowmax          = $_[1];        # End row
    my $colmin          = $_[2];        # Start column
    my $colmax          = $_[3];        # end column


    my $header          = pack("vv", $record, $length);
    my $data            = pack("v",  $grbit);
    $data              .= pack("C",  $chKey);
    $data              .= pack("C",  $cch);
    $data              .= pack("v",  $cce);
    $data              .= pack("v",  $unknown01);
    $data              .= pack("v",  $ixals);
    $data              .= pack("C",  $unknown02);
    $data              .= pack("C",  $cchCustMenu);
    $data              .= pack("C",  $cchDescription);
    $data              .= pack("C",  $cchHelptopic);
    $data              .= pack("C",  $cchStatustext);
    $data              .= pack("C",  $rgch);
    $data              .= pack("C",  $unknown03);
    $data              .= pack("v",  $ext_ref);

    $data              .= pack("v",  $rowmin);
    $data              .= pack("v",  $rowmax);
    $data              .= pack("v",  $colmin);
    $data              .= pack("v",  $colmax);

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
    my $length          = 0x002a;       # Number of bytes to follow

    my $index           = shift;        # Sheet index
    my $type            = shift;
    my $ext_ref         = shift;        # TODO

    my $grbit           = 0x0020;       # Option flags
    my $chKey           = 0x00;         # Keyboard shortcut
    my $cch             = 0x01;         # Length of text name
    my $cce             = 0x001a;       # Length of text definition
    my $unknown01       = 0x0000;       #
    my $ixals           = $index +1;    # Sheet index
    my $unknown02       = 0x00;         #
    my $cchCustMenu     = 0x00;         # Length of cust menu text
    my $cchDescription  = 0x00;         # Length of description text
    my $cchHelptopic    = 0x00;         # Length of help topic text
    my $cchStatustext   = 0x00;         # Length of status bar text
    my $rgch            = $type;        # Built-in name type

    my $unknown03       = 0x29;
    my $unknown04       = 0x0017;
    my $unknown05       = 0x3b;

    my $rowmin          = $_[0];        # Start row
    my $rowmax          = $_[1];        # End row
    my $colmin          = $_[2];        # Start column
    my $colmax          = $_[3];        # end column


    my $header          = pack("vv", $record, $length);
    my $data            = pack("v",  $grbit);
    $data              .= pack("C",  $chKey);
    $data              .= pack("C",  $cch);
    $data              .= pack("v",  $cce);
    $data              .= pack("v",  $unknown01);
    $data              .= pack("v",  $ixals);
    $data              .= pack("C",  $unknown02);
    $data              .= pack("C",  $cchCustMenu);
    $data              .= pack("C",  $cchDescription);
    $data              .= pack("C",  $cchHelptopic);
    $data              .= pack("C",  $cchStatustext);
    $data              .= pack("C",  $rgch);

    # Column definition
    $data              .= pack("C",  $unknown03);
    $data              .= pack("v",  $unknown04);
    $data              .= pack("C",  $unknown05);
    $data              .= pack("v",  $ext_ref);
    $data              .= pack("v",  0x0000);
    $data              .= pack("v",  0xffff);
    $data              .= pack("v",  $colmin);
    $data              .= pack("v",  $colmax);

    # Row definition
    $data              .= pack("C",  $unknown05);
    $data              .= pack("v",  $ext_ref);
    $data              .= pack("v",  $rowmin);
    $data              .= pack("v",  $rowmax);
    $data              .= pack("v",  0x00);
    $data              .= pack("v",  0xff);
    # End of data
    $data              .= pack("C",  0x10);

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


###############################################################################
#
# _store_hideobj()
#
# Stores the HIDEOBJ biff record.
#
sub _store_hideobj {

    my $self            = shift;

    my $record          = 0x008D;               # Record identifier
    my $length          = 0x0002;               # Number of bytes to follow
    my $hide            = $self->{_hideobj};    # Option to hide objects

    my $header          = pack("vv", $record, $length);
    my $data            = pack("v",  $hide);

    $self->_append($header, $data);
}




###############################################################################
###############################################################################
###############################################################################



###############################################################################
#
# _calculate_extern_sizes()
#
# We need to calculate the space required by the SUPBOOK, EXTERNSHEET and NAME
# records so that it can be added to the BOUNDSHEET offsets.
#
sub _calculate_extern_sizes {

    my $self   = shift;


    my %ext_refs        = $self->{_parser}->get_ext_sheets();
    my $ext_ref_count   = scalar keys %ext_refs;
    my $length          = 0;
    my $index           = 0;

    foreach my $worksheet (@{$self->{_worksheets}}) {

        my $rowmin      = $worksheet->{_title_rowmin};
        my $colmin      = $worksheet->{_title_colmin};
        my $key         = "$index:$index";
        $index++;


        # Print area NAME records
        if (defined $worksheet->{_print_rowmin}) {
            $ext_refs{$key} = $ext_ref_count++ if not exists $ext_refs{$key};

            $length += 31 ;
        }


        # Print title  NAME records
        if (defined $rowmin and defined $colmin) {
            $ext_refs{$key} = $ext_ref_count++ if not exists $ext_refs{$key};

            $length += 46;
        }
        elsif (defined $rowmin or defined $colmin) {
            $ext_refs{$key} = $ext_ref_count++ if not exists $ext_refs{$key};

            $length += 31;
        }
        else {
            # TODO
        }


    }


    # TODO
    $self->{_ext_ref_count} = $ext_ref_count;
    $self->{_ext_refs}      = {%ext_refs};



    # If there are no external refs then we don't write, SUPBOOK, EXTERNSHEET
    # and NAME. Therefore the length is 0.

    return $length = 0 if $ext_ref_count == 0;



    # The SUPBOOK record is 8 bytes
    $length += 8;

    # The EXTERNSHEET record is 6 bytes + 6 bytes for each external ref
    $length += 6 * (1 + $ext_ref_count);

    return $length;
}


###############################################################################
#
# _calculate_shared_string_sizes()
#
# Handling of the SST continue blocks is complicated by the need to include an
# additional continuation byte depending on whether the string is split between
# blocks or whether it starts at the beginning of the block. (There are also
# additional complications that will arise later when/if Rich Strings are
# supported). As such we cannot use the simple CONTINUE mechanism provided by
# the _add_continue() method in BIFFwriter.pm. Thus we have to make two passes
# through the strings data. The first is to calculate the required block sizes
# and the second, in _store_shared_strings(), is to write the actual strings.
# The first pass through the data is also used to calculate the size of the SST
# and CONTINUE records for use in setting the BOUNDSHEET record offsets. The
# downside of this is that the same algorithm repeated in _store_shared_strings.
#
sub _calculate_shared_string_sizes {

    my $self    = shift;

    my @strings;
    $#strings = $self->{_str_unique} -1; # Pre-extend array

    while (my $key = each %{$self->{_str_table}}) {
        $strings[$self->{_str_table}->{$key}] = $key;
    }

    # The SST data could be very large, free some memory (maybe).
    $self->{_str_table} = undef;
    $self->{_str_array} = [@strings];


    # Iterate through the strings to calculate the CONTINUE block sizes.
    #
    # The SST blocks requires a specialised CONTINUE block, so we have to
    # ensure that the maximum data block size is less than the limit used by
    # _add_continue() in BIFFwriter.pm. For simplicity we use the same size
    # for the SST and CONTINUE records:
    #   8228 : Maximum Excel97 block size
    #     -4 : Length of block header
    #     -8 : Length of additional SST header information
    #     -8 : Arbitrary number to keep within _add_continue() limit
    # = 8208
    #
    my $continue_limit = 8208;
    my $block_length   = 0;
    my $written        = 0;
    my @block_sizes;
    my $continue       = 0;

    for my $string (@strings) {

        my $string_length = length $string;
        my $encoding      = unpack "xx C", $string;
        my $split_string  = 0;


        # Block length is the total length of the strings that will be
        # written out in a single SST or CONTINUE block.
        #
        $block_length += $string_length;


        # We can write the string if it doesn't cross a CONTINUE boundary
        if ($block_length < $continue_limit) {
            $written += $string_length;
            next;
        }


        # Deal with the cases where the next string to be written will exceed
        # the CONTINUE boundary. If the string is very long it may need to be
        # written in more than one CONTINUE record.
        #
        while ($block_length >= $continue_limit) {

            # We need to avoid the case where a string is continued in the first
            # n bytes that contain the string header information.
            #
            my $header_length   = 3; # Min string + header size -1
            my $space_remaining = $continue_limit -$written -$continue;


            # Unicode data should only be split on char (2 byte) boundaries.
            # Therefore, in some cases we need to reduce the amount of available
            # space by 1 byte to ensure the correct alignment.
            my $align = 0;

            # Only applies to Unicode strings
            if ($encoding == 1) {
                # Min string + header size -1
                $header_length = 4;

                if ($space_remaining > $header_length) {
                    # String contains 3 byte header => split on odd boundary
                    if (not $split_string and $space_remaining % 2 != 1) {
                        $space_remaining--;
                        $align = 1;
                    }
                    # Split section without header => split on even boundary
                    elsif ($split_string and $space_remaining % 2 == 1) {
                        $space_remaining--;
                        $align = 1;
                    }

                    $split_string = 1;
                }
            }


            if ($space_remaining > $header_length) {
                # Write as much as possible of the string in the current block
                $written      += $space_remaining;

                # Reduce the current block length by the amount written
                $block_length -= $continue_limit -$continue -$align;

                # Store the max size for this block
                push @block_sizes, $continue_limit -$align;

                # If the current string was split then the next CONTINUE block
                # should have the string continue flag (grbit) set unless the
                # split string fits exactly into the remaining space.
                #
                if ($block_length > 0) {
                    $continue = 1;
                }
                else {
                    $continue = 0;
                }

            }
            else {
                # Store the max size for this block
                push @block_sizes, $written +$continue;

                # Not enough space to start the string in the current block
                $block_length -= $continue_limit -$space_remaining -$continue;
                $continue = 0;

            }

            # If the string (or substr) is small enough we can write it in the
            # new CONTINUE block. Else, go through the loop again to write it in
            # one or more CONTINUE blocks
            #
            if ($block_length < $continue_limit) {
                $written = $block_length;
            }
            else {
                $written = 0;
            }
        }
    }

    # Store the max size for the last block unless it is empty
    push @block_sizes, $written +$continue if $written +$continue;


    $self->{_str_block_sizes} = [@block_sizes];


    # Calculate the total length of the SST and associated CONTINUEs (if any).
    # The SST record will have a length even if it contains no strings.
    # This length is required to set the offsets in the BOUNDSHEET records since
    # they must be written before the SST records
    #
    my $length  = 12;
    $length    +=     shift @block_sizes if    @block_sizes; # SST
    $length    += 4 + shift @block_sizes while @block_sizes; # CONTINUEs

    return $length;
}


###############################################################################
#
# _store_shared_strings()
#
# Write all of the workbooks strings into an indexed array.
#
# See the comments in _calculate_shared_string_sizes() for more information.
#
# The Excel documentation says that the SST record should be followed by an
# EXTSST record. The EXTSST record is a hash table that is used to optimise
# access to SST. However, despite the documentation it doesn't seem to be
# required so we will ignore it.
#
sub _store_shared_strings {

    my $self                = shift;

    my @strings = @{$self->{_str_array}};


    my $record              = 0x00FC;   # Record identifier
    my $length              = 0x0008;   # Number of bytes to follow
    my $total               = 0x0000;

    # Iterate through the strings to calculate the CONTINUE block sizes
    my $continue_limit = 8208;
    my $block_length   = 0;
    my $written        = 0;
    my $continue       = 0;

    # The SST and CONTINUE block sizes have been pre-calculated by
    # _calculate_shared_string_sizes()
    my @block_sizes    = @{$self->{_str_block_sizes}};


    # The SST record is required even if it contains no strings. Thus we will
    # always have a length
    #
    if (@block_sizes) {
        $length = 8 + shift @block_sizes;
    }
    else {
        # No strings
        $length = 8;
    }

    # Write the SST block header information
    my $header      = pack("vv", $record, $length);
    my $data        = pack("VV", $self->{_str_total}, $self->{_str_unique});
    $self->_append($header, $data);


    # Iterate through the strings and write them out
    for my $string (@strings) {

        my $string_length = length $string;
        my $encoding      = unpack "xx C", $string;
        my $split_string  = 0;


        # Block length is the total length of the strings that will be
        # written out in a single SST or CONTINUE block.
        #
        $block_length += $string_length;


        # We can write the string if it doesn't cross a CONTINUE boundary
        if ($block_length < $continue_limit) {
            $self->_append($string);
            $written += $string_length;
            next;
        }


        # Deal with the cases where the next string to be written will exceed
        # the CONTINUE boundary. If the string is very long it may need to be
        # written in more than one CONTINUE record.
        #
        while ($block_length >= $continue_limit) {

            # We need to avoid the case where a string is continued in the first
            # n bytes that contain the string header information.
            #
            my $header_length   = 3; # Min string + header size -1
            my $space_remaining = $continue_limit -$written -$continue;


            # Unicode data should only be split on char (2 byte) boundaries.
            # Therefore, in some cases we need to reduce the amount of available
            # space by 1 byte to ensure the correct alignment.
            my $align = 0;

            # Only applies to Unicode strings
            if ($encoding == 1) {
                # Min string + header size -1
                $header_length = 4;

                if ($space_remaining > $header_length) {
                    # String contains 3 byte header => split on odd boundary
                    if (not $split_string and $space_remaining % 2 != 1) {
                        $space_remaining--;
                        $align = 1;
                    }
                    # Split section without header => split on even boundary
                    elsif ($split_string and $space_remaining % 2 == 1) {
                        $space_remaining--;
                        $align = 1;
                    }

                    $split_string = 1;
                }
            }


            if ($space_remaining > $header_length) {
                # Write as much as possible of the string in the current block
                my $tmp = substr $string, 0, $space_remaining;
                $self->_append($tmp);

                # The remainder will be written in the next block(s)
                $string = substr $string, $space_remaining;

                # Reduce the current block length by the amount written
                $block_length -= $continue_limit -$continue -$align;

                # If the current string was split then the next CONTINUE block
                # should have the string continue flag (grbit) set unless the
                # split string fits exactly into the remaining space.
                #
                if ($block_length > 0) {
                    $continue = 1;
                }
                else {
                    $continue = 0;
                }
            }
            else {
                # Not enough space to start the string in the current block
                $block_length -= $continue_limit -$space_remaining -$continue;
                $continue = 0;
            }

            # Write the CONTINUE block header
            if (@block_sizes) {
                $record  = 0x003C;
                $length  = shift @block_sizes;

                $header  = pack("vv", $record, $length);
                $header .= pack("C", $encoding) if $continue;

                $self->_append($header);
            }

            # If the string (or substr) is small enough we can write it in the
            # new CONTINUE block. Else, go through the loop again to write it in
            # one or more CONTINUE blocks
            #
            if ($block_length < $continue_limit) {
                $self->_append($string);

                $written = $block_length;
            }
            else {
                $written = 0;
            }
        }
    }
}

#
# Methods related to comments and MSO objects.
#

###############################################################################
#
# _add_mso_drawing_group()
#
# Write the MSODRAWINGGROUP record that keeps track of the Escher drawing
# objects in the file such as images, comments and filters.
#
sub _add_mso_drawing_group {

    my $self    = shift;

    return unless $self->{_mso_size};

    my $record  = 0x00EB;               # Record identifier
    my $length  = 0x0000;               # Number of bytes to follow

    my $data    = $self->_store_mso_dgg_container();
       $data   .= $self->_store_mso_dgg(@{$self->{_mso_clusters}});
       $data   .= $self->_store_mso_opt();
       $data   .= $self->_store_mso_split_menu_colors();

       $length  = length $data;
    my $header  = pack("vv", $record, $length);

    $self->_append($header, $data);


    return $header . $data;

}


###############################################################################
#
# _store_mso_dgg_container()
#
# Write the Escher DggContainer record that is part of MSODRAWINGGROUP.
#
sub _store_mso_dgg_container {

    my $self        = shift;

    my $type        = 0xF000;
    my $version     = 15;
    my $instance    = 0;
    my $data        = '';
    my $length      = $self->{_mso_size} -12; # -4 (biff header) -8 (for this).


    return $self->_add_mso_generic($type, $version, $instance, $data, $length);
}


###############################################################################
#
# _store_mso_dgg()
#
# Write the Escher Dgg record that is part of MSODRAWINGGROUP.
#
sub _store_mso_dgg {

    my $self        = shift;

    my $type            = 0xF006;
    my $version         = 0;
    my $instance        = 0;
    my $data            = '';
    my $length          = undef; # Calculate automatically.

    my $max_spid        = $_[0];
    my $num_clusters    = $_[1];
    my $shapes_saved    = $_[2];
    my $drawings_saved  = $_[3];
    my $clusters        = $_[4];

    $data               = pack "VVVV",  $max_spid,     $num_clusters,
                                        $shapes_saved, $drawings_saved;

    for my $aref (@$clusters) {
        my $drawing_id      = $aref->[0];
        my $shape_ids_used  = $aref->[1];

        $data              .= pack "VV",  $drawing_id, $shape_ids_used;
    }


    return $self->_add_mso_generic($type, $version, $instance, $data, $length);
}


###############################################################################
#
# _store_mso_opt()
#
# Write the Escher Opt record that is part of MSODRAWINGGROUP.
#
sub _store_mso_opt {

    my $self        = shift;

    my $type        = 0xF00B;
    my $version     = 3;
    my $instance    = 3;
    my $data        = '';
    my $length      = 18;

    $data           = pack "H*", 'BF0008000800810109000008C0014000' .
                                 '0008';


    return $self->_add_mso_generic($type, $version, $instance, $data, $length);
}


###############################################################################
#
# _store_mso_split_menu_colors()
#
# Write the Escher SplitMenuColors record that is part of MSODRAWINGGROUP.
#
sub _store_mso_split_menu_colors {

    my $self        = shift;

    my $type        = 0xF11E;
    my $version     = 0;
    my $instance    = 4;
    my $data        = '';
    my $length      = 16;

    $data           = pack "H*", '0D0000080C00000817000008F7000010';


    return $self->_add_mso_generic($type, $version, $instance, $data, $length);
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

© MM-MMVI, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
