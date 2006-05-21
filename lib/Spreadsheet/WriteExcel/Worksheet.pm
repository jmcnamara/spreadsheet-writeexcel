package Spreadsheet::WriteExcel::Worksheet;

###############################################################################
#
# Worksheet - A writer class for Excel Worksheets.
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
use Spreadsheet::WriteExcel::Format;
use Spreadsheet::WriteExcel::Formula;



use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::WriteExcel::BIFFwriter);

$VERSION = '2.16';

###############################################################################
#
# new()
#
# Constructor. Creates a new Worksheet object from a BIFFwriter object
#
sub new {

    my $class                       = shift;
    my $self                        = Spreadsheet::WriteExcel::BIFFwriter->new();
    my $rowmax                      = 65536;
    my $colmax                      = 256;
    my $strmax                      = 0;

    $self->{_name}                  = $_[0];
    $self->{_index}                 = $_[1];
    $self->{_encoding}              = $_[2];
    $self->{_activesheet}           = $_[3];
    $self->{_firstsheet}            = $_[4];
    $self->{_url_format}            = $_[5];
    $self->{_parser}                = $_[6];
    $self->{_tempdir}               = $_[7];

    $self->{_str_total}             = $_[8];
    $self->{_str_unique}            = $_[9];
    $self->{_str_table}             = $_[10];
    $self->{_1904}                  = $_[11];

    $self->{_type}                  = 0x0000;
    $self->{_ext_sheets}            = [];
    $self->{_using_tmpfile}         = 1;
    $self->{_filehandle}            = "";
    $self->{_fileclosed}            = 0;
    $self->{_offset}                = 0;
    $self->{_xls_rowmax}            = $rowmax;
    $self->{_xls_colmax}            = $colmax;
    $self->{_xls_strmax}            = $strmax;
    $self->{_dim_rowmin}            = $rowmax +1;
    $self->{_dim_rowmax}            = 0;
    $self->{_dim_colmin}            = $colmax +1;
    $self->{_dim_colmax}            = 0;
    $self->{_dim_changed}           = 0;
    $self->{_colinfo}               = [];
    $self->{_selection}             = [0, 0];
    $self->{_panes}                 = [];
    $self->{_active_pane}           = 3;
    $self->{_frozen}                = 0;
    $self->{_selected}              = 0;
    $self->{_hidden}                = 0;
    $self->{_active}                = 0;
    $self->{_tab_color}             = 0;

    $self->{_first_row}             = 0;
    $self->{_first_col}             = 0;
    $self->{_display_formulas}      = 0;
    $self->{_display_headers}       = 1;
    $self->{_display_zeros}         = 1;
    $self->{_display_arabic}        = 0;

    $self->{_paper_size}            = 0x0;
    $self->{_orientation}           = 0x1;
    $self->{_header}                = '';
    $self->{_footer}                = '';
    $self->{_header_encoding}       = 0;
    $self->{_footer_encoding}       = 0;
    $self->{_hcenter}               = 0;
    $self->{_vcenter}               = 0;
    $self->{_margin_header}         = 0.50;
    $self->{_margin_footer}         = 0.50;
    $self->{_margin_left}           = 0.75;
    $self->{_margin_right}          = 0.75;
    $self->{_margin_top}            = 1.00;
    $self->{_margin_bottom}         = 1.00;

    $self->{_title_rowmin}          = undef;
    $self->{_title_rowmax}          = undef;
    $self->{_title_colmin}          = undef;
    $self->{_title_colmax}          = undef;
    $self->{_print_rowmin}          = undef;
    $self->{_print_rowmax}          = undef;
    $self->{_print_colmin}          = undef;
    $self->{_print_colmax}          = undef;

    $self->{_print_gridlines}       = 1;
    $self->{_screen_gridlines}      = 1;
    $self->{_print_headers}         = 0;

    $self->{_page_order}            = 0;
    $self->{_black_white}           = 0;
    $self->{_draft_quality}         = 0;
    $self->{_print_comments}        = 0;
    $self->{_page_start}            = 1;

    $self->{_fit_page}              = 0;
    $self->{_fit_width}             = 0;
    $self->{_fit_height}            = 0;

    $self->{_hbreaks}               = [];
    $self->{_vbreaks}               = [];

    $self->{_protect}               = 0;
    $self->{_password}              = undef;

    $self->{_col_sizes}             = {};
    $self->{_row_sizes}             = {};

    $self->{_col_formats}           = {};
    $self->{_row_formats}           = {};

    $self->{_zoom}                  = 100;
    $self->{_print_scale}           = 100;
    $self->{_page_view}             = 0;

    $self->{_leading_zeros}         = 0;

    $self->{_outline_row_level}     = 0;
    $self->{_outline_style}         = 0;
    $self->{_outline_below}         = 1;
    $self->{_outline_right}         = 1;
    $self->{_outline_on}            = 1;

    $self->{_write_match}           = [];
    $self->{_comments}              = {};
    $self->{_comments_ids}          = [];
    $self->{_comments_author}       = '';
    $self->{_comments_author_enc}   = 0;
    $self->{_comments_visible}      = 0;

    $self->{_writing_url}           = 0;


    bless $self, $class;
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

        if ($self->{_index} == 0 && $^W) {
            my $dir = $self->{_tempdir} || File::Spec->tmpdir();

            warn "Unable to create temp files in $tmp_dir. Data will be ".
                 "stored in memory. Refer to set_tempdir() in the ".
                 "Spreadsheet::WriteExcel documentation.\n" ;
        }
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
    my $sheetnames = shift;
    my $num_sheets = scalar @$sheetnames;

    ################################################
    # Prepend in reverse order!!
    #

    # Prepend the sheet dimensions
    $self->_store_dimensions();

    # Prepend the COLINFO records if they exist
    if (@{$self->{_colinfo}}){
        my @colinfo = @{$self->{_colinfo}};
        while (@colinfo) {
            my $arrayref = pop @colinfo;
            $self->_store_colinfo(@$arrayref);
        }
    }

    # Add the DEFCOLWIDTH record
    $self->_store_defcol();

    # Prepend the sheet password
    $self->_store_password();

    # Prepend the sheet protection
    $self->_store_protect();
    $self->_store_obj_protect();

    # Prepend the page setup
    $self->_store_setup();

    # Prepend the bottom margin
    $self->_store_margin_bottom();

    # Prepend the top margin
    $self->_store_margin_top();

    # Prepend the right margin
    $self->_store_margin_right();

    # Prepend the left margin
    $self->_store_margin_left();

    # Prepend the page vertical centering
    $self->_store_vcenter();

    # Prepend the page horizontal centering
    $self->_store_hcenter();

    # Prepend the page footer
    $self->_store_footer();

    # Prepend the page header
    $self->_store_header();

    # Prepend the vertical page breaks
    $self->_store_vbreak();

    # Prepend the horizontal page breaks
    $self->_store_hbreak();

    # Prepend WSBOOL
    $self->_store_wsbool();

    # Prepend the default row height.
    $self->_store_defrow();

    # Prepend GUTS
    $self->_store_guts();

    # Prepend GRIDSET
    $self->_store_gridset();

    # Prepend PRINTGRIDLINES
    $self->_store_print_gridlines();

    # Prepend PRINTHEADERS
    $self->_store_print_headers();

    # Prepend the BOF record
    $self->_store_bof(0x0010);

    #
    # End of prepend. Read upwards from here.
    ################################################

    # Append
    $self->_store_comments();
    $self->_store_window2();
    $self->_store_page_view();
    $self->_store_zoom();
    $self->_store_panes(@{$self->{_panes}}) if @{$self->{_panes}};
    $self->_store_selection(@{$self->{_selection}});
    $self->_store_tab_color();
    $self->_store_eof();
}


###############################################################################
#
# get_name().
#
# Retrieve the worksheet name.
#
sub get_name {

    my $self    = shift;

    return $self->{_name};
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
# select()
#
# Set this worksheet as a selected worksheet, i.e. the worksheet has its tab
# highlighted.
#
sub select {

    my $self = shift;

    $self->{_hidden}         = 0; # Selected worksheet can't be hidden.
    $self->{_selected}       = 1;
}


###############################################################################
#
# activate()
#
# Set this worksheet as the active worksheet, i.e. the worksheet that is
# displayed when the workbook is opened. Also set it as selected.
#
sub activate {

    my $self = shift;

    $self->{_hidden}         = 0; # Active worksheet can't be hidden.
    $self->{_selected}       = 1;
    ${$self->{_activesheet}} = $self->{_index};
}


###############################################################################
#
# hide()
#
# Hide this worksheet.
#
sub hide {

    my $self = shift;

    $self->{_hidden}         = 1;

    # A hidden worksheet shouldn't be active or selected.
    $self->{_selected}       = 0;
    ${$self->{_activesheet}} = 0;
    ${$self->{_firstsheet}}  = 0;
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

    $self->{_hidden}         = 0; # Active worksheet can't be hidden.
    ${$self->{_firstsheet}}  = $self->{_index};
}


###############################################################################
#
# protect($password)
#
# Set the worksheet protection flag to prevent accidental modification and to
# hide formulas if the locked and hidden format properties have been set.
#
sub protect {

    my $self = shift;

    $self->{_protect}   = 1;
    $self->{_password}  = $self->_encode_password($_[0]) if defined $_[0];

}


###############################################################################
#
# set_column($firstcol, $lastcol, $width, $format, $hidden, $level)
#
# Set the width of a single column or a range of columns.
# See also: _store_colinfo
#
sub set_column {

    my $self = shift;
    my $cell = $_[0];

    # Check for a cell reference in A1 notation and substitute row and column
    if ($cell =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);

        # Returned values $row1 and $row2 aren't required here. Remove them.
        shift  @_;       # $row1
        splice @_, 1, 1; # $row2
    }

    return if @_ < 3; # Ensure at least $firstcol, $lastcol and $width
    return if not defined $_[0]; # Columns must be defined.
    return if not defined $_[1];

    push @{$self->{_colinfo}}, [ @_ ];


    # Store the col sizes for use when calculating image vertices taking
    # hidden columns into account. Also store the column formats.
    #
    my $width  = $_[4] ? 0 : $_[2]; # Set width to zero if column is hidden
       $width  ||= 0;                 # Ensure width isn't undef.
    my $format = $_[3];

    my ($firstcol, $lastcol) = @_;

    foreach my $col ($firstcol .. $lastcol) {
        $self->{_col_sizes}->{$col}   = $width;
        $self->{_col_formats}->{$col} = $format if defined $format;
    }
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

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    $self->{_selection} = [ @_ ];
}


###############################################################################
#
# freeze_panes()
#
# Set panes and mark them as frozen. See also _store_panes().
#
sub freeze_panes {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    $self->{_frozen} = 1;
    $self->{_panes}  = [ @_ ];
}


###############################################################################
#
# thaw_panes()
#
# Set panes and mark them as unfrozen. See also _store_panes().
#
sub thaw_panes {

    my $self = shift;

    $self->{_frozen} = 0;
    $self->{_panes}  = [ @_ ];
}


###############################################################################
#
# set_portrait()
#
# Set the page orientation as portrait.
#
sub set_portrait {

    my $self = shift;

    $self->{_orientation} = 1;
}


###############################################################################
#
# set_landscape()
#
# Set the page orientation as landscape.
#
sub set_landscape {

    my $self = shift;

    $self->{_orientation} = 0;
}


###############################################################################
#
# set_page_view()
#
# Set the page view mode for Mac Excel.
#
sub set_page_view {

    my $self = shift;

    $self->{_page_view} = defined $_[0] ? $_[0] : 1;
}


###############################################################################
#
# set_tab_color()
#
# Set the colour of the worksheet colour.
#
sub set_tab_color {

    my $self  = shift;

    my $color = &Spreadsheet::WriteExcel::Format::_get_color($_[0]);
       $color = 0 if $color == 0x7FFF; # Default color.

    $self->{_tab_color} = $color;
}


###############################################################################
#
# set_paper()
#
# Set the paper type. Ex. 1 = US Letter, 9 = A4
#
sub set_paper {

    my $self = shift;

    $self->{_paper_size} = $_[0] || 0;
}


###############################################################################
#
# set_header()
#
# Set the page header caption and optional margin.
#
sub set_header {

    my $self     = shift;
    my $string   = $_[0] || '';
    my $margin   = $_[1] || 0.50;
    my $encoding = $_[2] || 0;

    # Handle utf8 strings in perl 5.8.
    if ($] >= 5.008) {
        require Encode;

        if (Encode::is_utf8($string)) {
            $string = Encode::encode("UTF-16BE", $string);
            $encoding = 1;
        }
    }

    my $limit    = $encoding ? 255 *2 : 255;

    if (length $string >= $limit) {
        carp 'Header string must be less than 255 characters';
        return;
    }

    $self->{_header}          = $string;
    $self->{_margin_header}   = $margin;
    $self->{_header_encoding} = $encoding;
}


###############################################################################
#
# set_footer()
#
# Set the page footer caption and optional margin.
#
sub set_footer {

    my $self     = shift;
    my $string   = $_[0] || '';
    my $margin   = $_[1] || 0.50;
    my $encoding = $_[2] || 0;

    # Handle utf8 strings in perl 5.8.
    if ($] >= 5.008) {
        require Encode;

        if (Encode::is_utf8($string)) {
            $string = Encode::encode("UTF-16BE", $string);
            $encoding = 1;
        }
    }

    my $limit    = $encoding ? 255 *2 : 255;


    if (length $string >= $limit) {
        carp 'Footer string must be less than 255 characters';
        return;
    }

    $self->{_footer}          = $string;
    $self->{_margin_footer}   = $margin;
    $self->{_footer_encoding} = $encoding;
}


###############################################################################
#
# center_horizontally()
#
# Center the page horizontally.
#
sub center_horizontally {

    my $self = shift;

    if (defined $_[0]) {
        $self->{_hcenter} = $_[0];
    }
    else {
        $self->{_hcenter} = 1;
    }
}


###############################################################################
#
# center_vertically()
#
# Center the page horinzontally.
#
sub center_vertically {

    my $self = shift;

    if (defined $_[0]) {
        $self->{_vcenter} = $_[0];
    }
    else {
        $self->{_vcenter} = 1;
    }
}


###############################################################################
#
# set_margins()
#
# Set all the page margins to the same value in inches.
#
sub set_margins {

    my $self = shift;

    $self->set_margin_left($_[0]);
    $self->set_margin_right($_[0]);
    $self->set_margin_top($_[0]);
    $self->set_margin_bottom($_[0]);
}


###############################################################################
#
# set_margins_LR()
#
# Set the left and right margins to the same value in inches.
#
sub set_margins_LR {

    my $self = shift;

    $self->set_margin_left($_[0]);
    $self->set_margin_right($_[0]);
}


###############################################################################
#
# set_margins_TB()
#
# Set the top and bottom margins to the same value in inches.
#
sub set_margins_TB {

    my $self = shift;

    $self->set_margin_top($_[0]);
    $self->set_margin_bottom($_[0]);
}


###############################################################################
#
# set_margin_left()
#
# Set the left margin in inches.
#
sub set_margin_left {

    my $self = shift;

    $self->{_margin_left} = defined $_[0] ? $_[0] : 0.75;
}


###############################################################################
#
# set_margin_right()
#
# Set the right margin in inches.
#
sub set_margin_right {

    my $self = shift;

    $self->{_margin_right} = defined $_[0] ? $_[0] : 0.75;
}


###############################################################################
#
# set_margin_top()
#
# Set the top margin in inches.
#
sub set_margin_top {

    my $self = shift;

    $self->{_margin_top} = defined $_[0] ? $_[0] : 1.00;
}


###############################################################################
#
# set_margin_bottom()
#
# Set the bottom margin in inches.
#
sub set_margin_bottom {

    my $self = shift;

    $self->{_margin_bottom} = defined $_[0] ? $_[0] : 1.00;
}


###############################################################################
#
# repeat_rows($first_row, $last_row)
#
# Set the rows to repeat at the top of each printed page. See also the
# _store_name_xxxx() methods in Workbook.pm.
#
sub repeat_rows {

    my $self = shift;

    $self->{_title_rowmin}  = $_[0];
    $self->{_title_rowmax}  = $_[1] || $_[0]; # Second row is optional
}


###############################################################################
#
# repeat_columns($first_col, $last_col)
#
# Set the columns to repeat at the left hand side of each printed page.
# See also the _store_names() methods in Workbook.pm.
#
sub repeat_columns {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);

        # Returned values $row1 and $row2 aren't required here. Remove them.
        shift  @_;       # $row1
        splice @_, 1, 1; # $row2
    }

    $self->{_title_colmin}  = $_[0];
    $self->{_title_colmax}  = $_[1] || $_[0]; # Second col is optional
}


###############################################################################
#
# print_area($first_row, $first_col, $last_row, $last_col)
#
# Set the area of each worksheet that will be printed. See also the
# _store_names() methods in Workbook.pm.
#
sub print_area {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    return if @_ != 4; # Require 4 parameters

    $self->{_print_rowmin} = $_[0];
    $self->{_print_colmin} = $_[1];
    $self->{_print_rowmax} = $_[2];
    $self->{_print_colmax} = $_[3];
}


###############################################################################
#
# hide_gridlines()
#
# Set the option to hide gridlines on the screen and the printed page.
# There are two ways of doing this in the Excel BIFF format: The first is by
# setting the DspGrid field of the WINDOW2 record, this turns off the screen
# and subsequently the print gridline. The second method is to via the
# PRINTGRIDLINES and GRIDSET records, this turns off the printed gridlines
# only. The first method is probably sufficient for most cases. The second
# method is supported for backwards compatibility. Porters take note.
#
sub hide_gridlines {

    my $self   = shift;
    my $option = $_[0];

    $option = 1 unless defined $option; # Default to hiding printed gridlines

    if ($option == 0) {
        $self->{_print_gridlines}  = 1; # 1 = display, 0 = hide
        $self->{_screen_gridlines} = 1;
    }
    elsif ($option == 1) {
        $self->{_print_gridlines}  = 0;
        $self->{_screen_gridlines} = 1;
    }
    else {
        $self->{_print_gridlines}  = 0;
        $self->{_screen_gridlines} = 0;
    }
}


###############################################################################
#
# print_row_col_headers()
#
# Set the option to print the row and column headers on the printed page.
# See also the _store_print_headers() method below.
#
sub print_row_col_headers {

    my $self = shift;

    if (defined $_[0]) {
        $self->{_print_headers} = $_[0];
    }
    else {
        $self->{_print_headers} = 1;
    }
}


###############################################################################
#
# fit_to_pages($width, $height)
#
# Store the vertical and horizontal number of pages that will define the
# maximum area printed. See also _store_setup() and _store_wsbool() below.
#
sub fit_to_pages {

    my $self = shift;

    $self->{_fit_page}      = 1;
    $self->{_fit_width}     = $_[0] || 0;
    $self->{_fit_height}    = $_[1] || 0;
}


###############################################################################
#
# set_h_pagebreaks(@breaks)
#
# Store the horizontal page breaks on a worksheet.
#
sub set_h_pagebreaks {

    my $self = shift;

    push @{$self->{_hbreaks}}, @_;
}


###############################################################################
#
# set_v_pagebreaks(@breaks)
#
# Store the vertical page breaks on a worksheet.
#
sub set_v_pagebreaks {

    my $self = shift;

    push @{$self->{_vbreaks}}, @_;
}


###############################################################################
#
# set_zoom($scale)
#
# Set the worksheet zoom factor.
#
sub set_zoom {

    my $self  = shift;
    my $scale = $_[0] || 100;

    # Confine the scale to Excel's range
    if ($scale < 10 or $scale > 400) {
        carp "Zoom factor $scale outside range: 10 <= zoom <= 400";
        $scale = 100;
    }

    $self->{_zoom} = int $scale;
}


###############################################################################
#
# set_print_scale($scale)
#
# Set the scale factor for the printed page.
#
sub set_print_scale {

    my $self  = shift;
    my $scale = $_[0] || 100;

    # Confine the scale to Excel's range
    if ($scale < 10 or $scale > 400) {
        carp "Print scale $scale outside range: 10 <= zoom <= 400";
        $scale = 100;
    }

    # Turn off "fit to page" option
    $self->{_fit_page}    = 0;

    $self->{_print_scale} = int $scale;
}


###############################################################################
#
# keep_leading_zeros()
#
# Causes the write() method to treat integers with a leading zero as a string.
# This ensures that any leading zeros such, as in zip codes, are maintained.
#
sub keep_leading_zeros {

    my $self = shift;

    if (defined $_[0]) {
        $self->{_leading_zeros} = $_[0];
    }
    else {
        $self->{_leading_zeros} = 1;
    }
}


###############################################################################
#
# show_comments()
#
# Make any comments in the worksheet visible.
#
sub show_comments {

    my $self = shift;

    $self->{_comments_visible} = defined $_[0] ? $_[0] : 1;
}


###############################################################################
#
# set_comments_author()
#
# Set the default author of the cell comments.
#
sub set_comments_author {

    my $self = shift;

    $self->{_comments_author}     = defined $_[0] ? $_[0] : '';
    $self->{_comments_author_enc} =         $_[1] ? 1     : 0;
}


###############################################################################
#
# right_to_left()
#
# Display the worksheet right to left for some eastern versions of Excel.
#
sub right_to_left {

    my $self = shift;

    $self->{_display_arabic} = defined $_[0] ? $_[0] : 1;
}


###############################################################################
#
# hide_zero()
#
# Hide cell zero values.
#
sub hide_zero {

    my $self = shift;

    $self->{_display_zeros} = defined $_[0] ? not $_[0] : 0;
}


###############################################################################
#
# print_across()
#
# Set the order in which pages are printed.
#
sub print_across {

    my $self = shift;

    $self->{_page_order} = defined $_[0] ? $_[0] : 1;
}


###############################################################################
#
# set_first_row_column()
#
# Set the topmost and leftmost visible row and column.
# TODO: Document this when tested fully for interaction with panes.
#
sub set_first_row_column {

    my $self = shift;

    my $row  = $_[0] || 0;
    my $col  = $_[1] || 0;

    $row = 65535 if $row > 65535;
    $col = 255   if $col > 255;

    $self->{_first_row} = $row;
    $self->{_first_col} = $col;
}


###############################################################################
#
# add_write_handler($re, $code_ref)
#
# Allow the user to add their own matches and handlers to the write() method.
#
sub add_write_handler {

    my $self = shift;

    return unless @_ == 2;
    return unless ref $_[1] eq 'CODE';

    push @{$self->{_write_match}}, [ @_ ];
}



###############################################################################
#
# write($row, $col, $token, $format)
#
# Parse $token and call appropriate write method. $row and $column are zero
# indexed. $format is optional.
#
# The write_url() methods have a flag to prevent recursion when writing a
# string that looks like a url.
#
# Returns: return value of called subroutine
#
sub write {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    my $token = $_[2];

    # Handle undefs as blanks
    $token = '' unless defined $token;


    # First try user defined matches.
    for my $aref (@{$self->{_write_match}}) {
        my $re  = $aref->[0];
        my $sub = $aref->[1];

        if ($token =~ /$re/) {
            my $match = &$sub($self, @_);
            return $match if defined $match;
        }
    }


    # Match an array ref.
    if (ref $token eq "ARRAY") {
        return $self->write_row(@_);
    }
    # Match integer with leading zero(s)
    elsif ($self->{_leading_zeros} and $token =~ /^0\d+$/) {
        return $self->write_string(@_);
    }
    # Match number
    elsif ($token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/) {
        return $self->write_number(@_);
    }
    # Match http, https or ftp URL
    elsif ($token =~ m|^[fh]tt?ps?://|    and not $self->{_writing_url}) {
        return $self->write_url(@_);
    }
    # Match mailto:
    elsif ($token =~ m/^mailto:/          and not $self->{_writing_url}) {
        return $self->write_url(@_);
    }
    # Match internal or external sheet link
    elsif ($token =~ m[^(?:in|ex)ternal:] and not $self->{_writing_url}) {
        return $self->write_url(@_);
    }
    # Match formula
    elsif ($token =~ /^=/) {
        return $self->write_formula(@_);
    }
    # Match blank
    elsif ($token eq '') {
        splice @_, 2, 1; # remove the empty string from the parameter list
        return $self->write_blank(@_);
    }
    else {
        return $self->write_string(@_);
    }
}


###############################################################################
#
# write_row($row, $col, $array_ref, $format)
#
# Write a row of data starting from ($row, $col). Call write_col() if any of
# the elements of the array ref are in turn array refs. This allows the writing
# of 1D or 2D arrays of data in one go.
#
# Returns: the first encountered error value or zero for no errors
#
sub write_row {

    my $self = shift;


    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    # Catch non array refs passed by user.
    if (ref $_[2] ne 'ARRAY') {
        croak "Not an array ref in call to write_row()$!";
    }

    my $row     = shift;
    my $col     = shift;
    my $tokens  = shift;
    my @options = @_;
    my $error   = 0;
    my $ret;

    foreach my $token (@$tokens) {

        # Check for nested arrays
        if (ref $token eq "ARRAY") {
            $ret = $self->write_col($row, $col, $token, @options);
        } else {
            $ret = $self->write    ($row, $col, $token, @options);
        }

        # Return only the first error encountered, if any.
        $error ||= $ret;
        $col++;
    }

    return $error;
}


###############################################################################
#
# write_col($row, $col, $array_ref, $format)
#
# Write a column of data starting from ($row, $col). Call write_row() if any of
# the elements of the array ref are in turn array refs. This allows the writing
# of 1D or 2D arrays of data in one go.
#
# Returns: the first encountered error value or zero for no errors
#
sub write_col {

    my $self = shift;


    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    # Catch non array refs passed by user.
    if (ref $_[2] ne 'ARRAY') {
        croak "Not an array ref in call to write_row()$!";
    }

    my $row     = shift;
    my $col     = shift;
    my $tokens  = shift;
    my @options = @_;
    my $error   = 0;
    my $ret;

    foreach my $token (@$tokens) {

        # write() will deal with any nested arrays
        $ret = $self->write($row, $col, $token, @options);

        # Return only the first error encountered, if any.
        $error ||= $ret;
        $row++;
    }

    return $error;
}


###############################################################################
#
# write_comment($row, $col, $comment)
#
# Write a comment to the specified row and column (zero indexed).
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub write_comment {

    my $self = shift;


    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    if (@_ < 3) { return -1 } # Check the number of args


    my $row = $_[0];
    my $col = $_[1];

    # Check for pairs of optional arguments, i.e. an odd number of args.
    croak "Uneven number of additional arguments" unless @_ % 2;


    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);


    # We have to avoid duplicate comments in cells or else Excel will complain.
    $self->{_comments}->{$row}->{$col} = [ $self->_comment_params(@_) ];

}


###############################################################################
#
# _XF()
#
# Returns an index to the XF record in the workbook.
#
# Note: this is a function, not a method.
#
sub _XF {

    my $self   = $_[0];
    my $row    = $_[1];
    my $col    = $_[2];
    my $format = $_[3];

    my $error = "Error: refer to merge_range() in the documentation. " .
                 "Can't use previously merged format in non-merged cell";

    if (ref($format)) {
        # Temp code to prevent merged formats in non-merged cells.
        croak $error if $format->{_used_merge} == 1;
        $format->{_used_merge} = -1;

        return $format->get_xf_index();
    }
    elsif (exists $self->{_row_formats}->{$row}) {
        # Temp code to prevent merged formats in non-merged cells.
        croak $error if $self->{_row_formats}->{$row}->{_used_merge} == 1;
        $self->{_row_formats}->{$row}->{_used_merge} = -1;

        return $self->{_row_formats}->{$row}->get_xf_index();
    }
    elsif (exists $self->{_col_formats}->{$col}) {
        # Temp code to prevent merged formats in non-merged cells.
        croak $error if $self->{_col_formats}->{$col}->{_used_merge} == 1;
        $self->{_col_formats}->{$col}->{_used_merge} = -1;

        return $self->{_col_formats}->{$col}->get_xf_index();
    }
    else {
        return 0x0F;
    }
}


###############################################################################
###############################################################################
#
# Internal methods
#


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
# _substitute_cellref()
#
# Substitute an Excel cell reference in A1 notation for  zero based row and
# column values in an argument list.
#
# Ex: ("A4", "Hello") is converted to (3, 0, "Hello").
#
sub _substitute_cellref {

    my $self = shift;
    my $cell = uc(shift);

    # Convert a column range: 'A:A' or 'B:G'.
    # A range such as A:A is equivalent to A1:65536, so add rows as required
    if ($cell =~ /\$?([A-I]?[A-Z]):\$?([A-I]?[A-Z])/) {
        my ($row1, $col1) =  $self->_cell_to_rowcol($1 .'1');
        my ($row2, $col2) =  $self->_cell_to_rowcol($2 .'65536');
        return $row1, $col1, $row2, $col2, @_;
    }

    # Convert a cell range: 'A1:B7'
    if ($cell =~ /\$?([A-I]?[A-Z]\$?\d+):\$?([A-I]?[A-Z]\$?\d+)/) {
        my ($row1, $col1) =  $self->_cell_to_rowcol($1);
        my ($row2, $col2) =  $self->_cell_to_rowcol($2);
        return $row1, $col1, $row2, $col2, @_;
    }

    # Convert a cell reference: 'A1' or 'AD2000'
    if ($cell =~ /\$?([A-I]?[A-Z]\$?\d+)/) {
        my ($row1, $col1) =  $self->_cell_to_rowcol($1);
        return $row1, $col1, @_;

    }

    croak("Unknown cell reference $cell");
}


###############################################################################
#
# _cell_to_rowcol($cell_ref)
#
# Convert an Excel cell reference in A1 notation to a zero based row and column
# reference; converts C1 to (0, 2).
#
# Returns: row, column
#
sub _cell_to_rowcol {

    my $self = shift;
    my $cell = shift;

    $cell =~ /\$?([A-I]?[A-Z])\$?(\d+)/;

    my $col     = $1;
    my $row     = $2;

    # Convert base26 column string to number
    # All your Base are belong to us.
    my @chars = split //, $col;
    my $expn  = 0;
    $col      = 0;

    while (@chars) {
        my $char = pop(@chars); # LS char first
        $col += (ord($char) -ord('A') +1) * (26**$expn);
        $expn++;
    }

    # Convert 1-index to zero-index
    $row--;
    $col--;

    return $row, $col;
}


###############################################################################
#
# _sort_pagebreaks()
#
#
# This is an internal method that is used to filter elements of the array of
# pagebreaks used in the _store_hbreak() and _store_vbreak() methods. It:
#   1. Removes duplicate entries from the list.
#   2. Sorts the list.
#   3. Removes 0 from the list if present.
#
sub _sort_pagebreaks {

    my $self= shift;

    my %hash;
    my @array;

    @hash{@_} = undef;                       # Hash slice to remove duplicates
    @array    = sort {$a <=> $b} keys %hash; # Numerical sort
    shift @array if $array[0] == 0;          # Remove zero

    # 1000 vertical pagebreaks appears to be an internal Excel 5 limit.
    # It is slightly higher in Excel 97/200, approx. 1026
    splice(@array, 1000) if (@array > 1000);

    return @array
}


###############################################################################
#
# _encode_password($password)
#
# Based on the algorithm provided by Daniel Rentz of OpenOffice.
#
#
sub _encode_password {

    use integer;

    my $self      = shift;
    my $plaintext = $_[0];
    my $password;
    my $count;
    my @chars;
    my $i = 0;

    $count = @chars = split //, $plaintext;

    foreach my $char (@chars) {
        my $low_15;
        my $high_15;
        $char     = ord($char) << ++$i;
        $low_15   = $char & 0x7fff;
        $high_15  = $char & 0x7fff << 15;
        $high_15  = $high_15 >> 15;
        $char     = $low_15 | $high_15;
    }

    $password  = 0x0000;
    $password ^= $_ for @chars;
    $password ^= $count;
    $password ^= 0xCE4B;

    return $password;
}


###############################################################################
#
# outline_settings($visible, $symbols_below, $symbols_right, $auto_style)
#
# This method sets the properties for outlining and grouping. The defaults
# correspond to Excel's defaults.
#
sub outline_settings {

    my $self                = shift;

    $self->{_outline_on}    = defined $_[0] ? $_[0] : 1;
    $self->{_outline_below} = defined $_[1] ? $_[1] : 1;
    $self->{_outline_right} = defined $_[2] ? $_[2] : 1;
    $self->{_outline_style} =         $_[3] || 0;

    # Ensure this is a boolean vale for Window2
    $self->{_outline_on}    = 1 if $self->{_outline_on};
}




###############################################################################
###############################################################################
#
# BIFF RECORDS
#


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

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    if (@_ < 3) { return -1 }                    # Check the number of args

    my $record  = 0x0203;                        # Record identifier
    my $length  = 0x000E;                        # Number of bytes to follow

    my $row     = $_[0];                         # Zero indexed row
    my $col     = $_[1];                         # Zero indexed column
    my $num     = $_[2];
    my $xf      = _XF($self, $row, $col, $_[3]); # The cell format

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);

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

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    if (@_ < 3) { return -1 }                        # Check the number of args

    my $record      = 0x00FD;                        # Record identifier
    my $length      = 0x000A;                        # Bytes to follow

    my $row         = $_[0];                         # Zero indexed row
    my $col         = $_[1];                         # Zero indexed column
    my $strlen      = length($_[2]);
    my $str         = $_[2];
    my $xf          = _XF($self, $row, $col, $_[3]); # The cell format
    my $encoding    = 0x0;
    my $str_error   = 0;


    # Handle utf8 strings in perl 5.8.
    if ($] >= 5.008) {
        require Encode;

        if (Encode::is_utf8($str)) {
            my $tmp = Encode::encode("UTF-16LE", $str);
            return $self->write_unicode_le($row, $col, $tmp, $_[3]);
        }
    }


    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);

    # Limit the string to the max number of chars.
    if ($strlen > 32767) {
        $str       = substr($str, 0, 32767);
        $str_error = -3;
    }


    # Preprend the string with the type.
    my $str_header  = pack("vC", length($str), $encoding);
    $str            = $str_header . $str;


    if (not exists ${$self->{_str_table}}->{$str}) {
        ${$self->{_str_table}}->{$str} = ${$self->{_str_unique}}++;
    }


    ${$self->{_str_total}}++;


    my $header = pack("vv",   $record, $length);
    my $data   = pack("vvvV", $row, $col, $xf, ${$self->{_str_table}}->{$str});

    $self->_append($header, $data);

    return $str_error;
}


###############################################################################
#
# write_blank($row, $col, $format)
#
# Write a blank cell to the specified row and column (zero indexed).
# A blank cell is used to specify formatting without adding a string
# or a number.
#
# A blank cell without a format serves no purpose. Therefore, we don't write
# a BLANK record unless a format is specified. This is mainly an optimisation
# for the write_row() and write_col() methods.
#
# Returns  0 : normal termination (including no format)
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub write_blank {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    # Check the number of args
    return -1 if @_ < 2;

    # Don't write a blank cell unless it has a format
    return 0 if not defined $_[2];


    my $record  = 0x0201;                        # Record identifier
    my $length  = 0x0006;                        # Number of bytes to follow

    my $row     = $_[0];                         # Zero indexed row
    my $col     = $_[1];                         # Zero indexed column
    my $xf      = _XF($self, $row, $col, $_[2]); # The cell format

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);

    my $header    = pack("vv",  $record, $length);
    my $data      = pack("vvv", $row, $col, $xf);

    $self->_append($header, $data);

    return 0;
}


###############################################################################
#
# write_formula($row, $col, $formula, $format, $value)
#
# Write a formula to the specified row and column (zero indexed).
# The textual representation of the formula is passed to the parser in
# Formula.pm which returns a packed binary string.
#
# $format is optional.
#
# $value is an optional result of the formula that can be supplied by the user.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub write_formula{

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    if (@_ < 3) { return -1 }   # Check the number of args

    my $record    = 0x0006;     # Record identifier
    my $length;                 # Bytes to follow

    my $row       = $_[0];      # Zero indexed row
    my $col       = $_[1];      # Zero indexed column
    my $formula   = $_[2];      # The formula text string
    my $value     = $_[4];      # The formula text string


    # Excel normally stores the last calculated value of the formula in $num.
    # Clearly we are not in a position to calculate this a priori. Instead
    # we set $num to zero and set the option flags in $grbit to ensure
    # automatic calculation of the formula when the file is opened.

    # As a workaround for some non-Excel apps we also allow the user to
    # specify the result of the formula.
    #
    my $xf        = _XF($self, $row, $col, $_[3]);  # The cell format
    my $chn       = 0x0000;                         # Must be zero
    my $is_string = 0;                              # Formula evaluates to str
    my $num;                                        # Current value of formula
    my $grbit;                                      # Option flags

    ($num, $grbit, $is_string) = $self->_encode_formula_result($value);


    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);

    # Strip the = sign at the beginning of the formula string
    $formula    =~ s(^=)();

    my $tmp     = $formula;

    # Parse the formula using the parser in Formula.pm
    my $parser  = $self->{_parser};

    # In order to raise formula errors from the point of view of the calling
    # program we use an eval block and re-raise the error from here.
    #
    eval { $formula = $parser->parse_formula($formula) };

    if ($@) {
        $@ =~ s/\n$//;  # Strip the \n used in the Formula.pm die()
        croak $@;       # Re-raise the error
    }


    my $formlen = length($formula); # Length of the binary string
       $length  = 0x16 + $formlen;  # Length of the record data

    my $header  = pack("vv",    $record, $length);
    my $data    = pack("vvv",   $row, $col, $xf);
       $data   .= $num;
       $data   .= pack("vVv",   $grbit, $chn, $formlen);

    $self->_append($header, $data, $formula);

    $self->_write_formula_string($value) if $is_string;

    return 0;
}
###############################################################################
#
# _encode_formula_result()
#
# Encode the user supplied result for a formula.
#
sub _encode_formula_result{

    my $self = shift;

    my $value     = $_[0];      # Result to be encoded.
    my $is_string = 0;          # Formula evaluates to str.
    my $num;                    # Current value of formula.
    my $grbit;                  # Option flags.

    if (not defined $value) {
        $grbit  = 0x03;
        $num    = pack "d", 0;
    }
    else {
        # The user specified the result of the formula. We turn off the recalc
        # flag and check the result type.
        $grbit  = 0x00;

        if ($value =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/) {
            # Value is a number.
            $num = pack "d", $value;
        }
        else {

            my %bools = (
                            'TRUE'    => [1,  1],
                            'FALSE'   => [1,  0],
                            '#NULL!'  => [2,  0],
                            '#DIV/0!' => [2,  7],
                            '#VALUE!' => [2, 15],
                            '#REF!'   => [2, 23],
                            '#NAME?'  => [2, 29],
                            '#NUM!'   => [2, 36],
                            '#N/A'    => [2, 42],
                        );

            if (exists $bools{$value}) {
                # Value is a boolean.
                $num = pack "vvvv", $bools{$value}->[0],
                                    $bools{$value}->[1],
                                    0,
                                    0xFFFF;
            }
            else {
                # Value is a string.
                $num = pack "vvvv", 0,
                                    0,
                                    0,
                                    0xFFFF;
                $is_string = 1;
            }
        }
    }

    return ($num, $grbit, $is_string);
}


###############################################################################
#
# _write_formula_string()
#
# Write the string value when a formula evaluates to a string. The value cannot
# be calculated by the module and thus must be supplied by the user.
#
sub _write_formula_string {

    my $self = shift;

    my $record    = 0x0207;         # Record identifier
    my $length    = 0x00;           # Bytes to follow
    my $string    = $_[0];          # Formula string.
    my $strlen    = length $_[0];   # Length of the formula string (chars).
    my $encoding  = 0;              # String encoding.


    # Handle utf8 strings in perl 5.8.
    if ($] >= 5.008) {
        require Encode;

        if (Encode::is_utf8($string)) {
            $string = Encode::encode("UTF-16BE", $string);
            $encoding = 1;
        }
    }


    $length       = 0x03 + length $string;  # Length of the record data

    my $header    = pack("vv", $record, $length);
    my $data      = pack("vC", $strlen, $encoding);

    $self->_append($header, $data, $string);

    return 0;
}


###############################################################################
#
# store_formula($formula)
#
# Pre-parse a formula. This is used in conjunction with repeat_formula()
# to repetitively rewrite a formula without re-parsing it.
#
sub store_formula{

    my $self    = shift;
    my $formula = $_[0];      # The formula text string

    # Strip the = sign at the beginning of the formula string
    $formula    =~ s(^=)();

    # Parse the formula using the parser in Formula.pm
    my $parser  = $self->{_parser};

    # In order to raise formula errors from the point of view of the calling
    # program we use an eval block and re-raise the error from here.
    #
    my @tokens;
    eval { @tokens = $parser->parse_formula($formula) };

    if ($@) {
        $@ =~ s/\n$//;  # Strip the \n used in the Formula.pm die()
        croak $@;       # Re-raise the error
    }


    # Return the parsed tokens in an anonymous array
    return [@tokens];
}


###############################################################################
#
# repeat_formula($row, $col, $formula, $format, ($pattern => $replacement,...))
#
# Write a formula to the specified row and column (zero indexed) by
# substituting $pattern $replacement pairs in the $formula created via
# store_formula(). This allows the user to repetitively rewrite a formula
# without the significant overhead of parsing.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub repeat_formula {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    if (@_ < 2) { return -1 }   # Check the number of args

    my $record      = 0x0006;   # Record identifier
    my $length;                 # Bytes to follow

    my $row         = shift;    # Zero indexed row
    my $col         = shift;    # Zero indexed column
    my $formula_ref = shift;    # Array ref with formula tokens
    my $format      = shift;    # XF format
    my @pairs       = @_;       # Pattern/replacement pairs


    # Enforce an even number of arguments in the pattern/replacement list
    croak "Odd number of elements in pattern/replacement list" if @pairs %2;

    # Check that $formula is an array ref
    croak "Not a valid formula" if ref $formula_ref ne 'ARRAY';

    my @tokens  = @$formula_ref;

    # Ensure that there are tokens to substitute
    croak "No tokens in formula" unless @tokens;


    # As a temporary and undocumented measure we allow the user to specify the
    # result of the formula by appending a result => $value pair to the end
    # of the arguments.
    my $value = undef;
    if ($pairs[-2] eq 'result') {
        $value = pop @pairs;
                 pop @pairs;
    }


    while (@pairs) {
        my $pattern = shift @pairs;
        my $replace = shift @pairs;

        foreach my $token (@tokens) {
            last if $token =~ s/$pattern/$replace/;
        }
    }


    # Change the parameters in the formula cached by the Formula.pm object
    my $parser    = $self->{_parser};
    my $formula   = $parser->parse_tokens(@tokens);

    croak "Unrecognised token in formula" unless defined $formula;


    # Excel normally stores the last calculated value of the formula in $num.
    # Clearly we are not in a position to calculate this a priori. Instead
    # we set $num to zero and set the option flags in $grbit to ensure
    # automatic calculation of the formula when the file is opened.

    # As a workaround for some non-Excel apps we also allow the user to
    # specify the result of the formula.
    #
    my $xf        = _XF($self, $row, $col, $format); # The cell format
    my $chn       = 0x0000;                          # Must be zero
    my $is_string = 0;                               # Formula evaluates to str
    my $num;                                         # Current value of formula
    my $grbit;                                       # Option flags

    ($num, $grbit, $is_string) = $self->_encode_formula_result($value);

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);


    my $formlen   = length($formula); # Length of the binary string
    $length       = 0x16 + $formlen;  # Length of the record data

    my $header    = pack("vv",    $record, $length);
    my $data      = pack("vvv",   $row, $col, $xf);
       $data     .= $num;
       $data     .= pack("vVv",   $grbit, $chn, $formlen);


    $self->_append($header, $data, $formula);

    $self->_write_formula_string($value) if $is_string;

    return 0;
}


###############################################################################
#
# write_url($row, $col, $url, $string, $format)
#
# Write a hyperlink. This is comprised of two elements: the visible label and
# the invisible link. The visible label is the same as the link unless an
# alternative string is specified.
#
# The parameters $string and $format are optional and their order is
# interchangeable for backward compatibity reasons.
#
# The hyperlink can be to a http, ftp, mail, internal sheet, or external
# directory url.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long string truncated to 255 chars
#
sub write_url {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    # Check the number of args
    return -1 if @_ < 3;

    # Add start row and col to arg list
    return $self->write_url_range($_[0], $_[1], @_);
}


###############################################################################
#
# write_url_range($row1, $col1, $row2, $col2, $url, $string, $format)
#
# This is the more general form of write_url(). It allows a hyperlink to be
# written to a range of cells. This function also decides the type of hyperlink
# to be written. These are either, Web (http, ftp, mailto), Internal
# (Sheet1!A1) or external ('c:\temp\foo.xls#Sheet1!A1').
#
# See also write_url() above for a general description and return values.
#
sub write_url_range {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    # Check the number of args
    return -1 if @_ < 5;


    # Reverse the order of $string and $format if necessary. We work on a copy
    # in order to protect the callers args. We don't use "local @_" in case of
    # perl50005 threads.
    #
    my @args = @_;

    ($args[5], $args[6]) = ($args[6], $args[5]) if ref $args[5];

    my $url = $args[4];


    # Check for internal/external sheet links or default to web link
    return $self->_write_url_internal(@args) if $url =~ m[^internal:];
    return $self->_write_url_external(@args) if $url =~ m[^external:];
    return $self->_write_url_web(@args);
}


###############################################################################
#
# _write_url_web($row1, $col1, $row2, $col2, $url, $string, $format)
#
# Used to write http, ftp and mailto hyperlinks.
# The link type ($options) is 0x03 is the same as absolute dir ref without
# sheet. However it is differentiated by the $unknown2 data stream.
#
# See also write_url() above for a general description and return values.
#
sub _write_url_web {

    my $self    = shift;

    my $record      = 0x01B8;                       # Record identifier
    my $length      = 0x00000;                      # Bytes to follow

    my $row1        = $_[0];                        # Start row
    my $col1        = $_[1];                        # Start column
    my $row2        = $_[2];                        # End row
    my $col2        = $_[3];                        # End column
    my $url         = $_[4];                        # URL string
    my $str         = $_[5];                        # Alternative label
    my $xf          = $_[6] || $self->{_url_format};# The cell format


    # Write the visible label but protect against url recursion in write().
    $str                  = $url unless defined $str;
    $self->{_writing_url} = 1;
    my $error             = $self->write($row1, $col1, $str, $xf);
    $self->{_writing_url} = 0;
    return $error         if $error == -2;


    # Pack the undocumented parts of the hyperlink stream
    my $unknown1    = pack("H*", "D0C9EA79F9BACE118C8200AA004BA90B02000000");
    my $unknown2    = pack("H*", "E0C9EA79F9BACE118C8200AA004BA90B");


    # Pack the option flags
    my $options     = pack("V", 0x03);

    # Convert URL to a null terminated wchar string
    $url            = join("\0", split('', $url));
    $url            = $url . "\0\0\0";


    # Pack the length of the URL
    my $url_len     = pack("V", length($url));


    # Calculate the data length
    $length         = 0x34 + length($url);


    # Pack the header data
    my $header      = pack("vv",   $record, $length);
    my $data        = pack("vvvv", $row1, $row2, $col1, $col2);


    # Write the packed data
    $self->_append( $header,
                    $data,
                    $unknown1,
                    $options,
                    $unknown2,
                    $url_len,
                    $url);

    return $error;
}


###############################################################################
#
# _write_url_internal($row1, $col1, $row2, $col2, $url, $string, $format)
#
# Used to write internal reference hyperlinks such as "Sheet1!A1".
#
# See also write_url() above for a general description and return values.
#
sub _write_url_internal {

    my $self    = shift;

    my $record      = 0x01B8;                       # Record identifier
    my $length      = 0x00000;                      # Bytes to follow

    my $row1        = $_[0];                        # Start row
    my $col1        = $_[1];                        # Start column
    my $row2        = $_[2];                        # End row
    my $col2        = $_[3];                        # End column
    my $url         = $_[4];                        # URL string
    my $str         = $_[5];                        # Alternative label
    my $xf          = $_[6] || $self->{_url_format};# The cell format

    # Strip URL type
    $url            =~ s[^internal:][];


    # Write the visible label but protect against url recursion in write().
    $str                  = $url unless defined $str;
    $self->{_writing_url} = 1;
    my $error             = $self->write($row1, $col1, $str, $xf);
    $self->{_writing_url} = 0;
    return $error         if $error == -2;


    # Pack the undocumented parts of the hyperlink stream
    my $unknown1    = pack("H*", "D0C9EA79F9BACE118C8200AA004BA90B02000000");


    # Pack the option flags
    my $options     = pack("V", 0x08);


    # URL encoding.
    my $encoding    = 0;


    # Convert an Utf8 URL type and to a null terminated wchar string.
    if ($] >= 5.008) {
        require Encode;

        if (Encode::is_utf8($url)) {
            # Quote sheet name if not already, i.e., Sheet!A1 to 'Sheet!A1'.
            $url      =~ s/^(.+)!/'$1'!/ if not $url =~ /^'/;

            $url      = Encode::encode("UTF-16LE", $url);
            $url     .= "\0\0"; # URL is null terminated.
            $encoding = 1;
        }
    }


    # Convert an Ascii URL type and to a null terminated wchar string.
    if ($encoding == 0) {
        $url       .= "\0";
        $url        = pack 'v*', unpack 'c*', $url;
    }


    # Pack the length of the URL as chars (not wchars)
    my $url_len     = pack("V", int(length($url)/2));


    # Calculate the data length
    $length         = 0x24 + length($url);


    # Pack the header data
    my $header      = pack("vv",   $record, $length);
    my $data        = pack("vvvv", $row1, $row2, $col1, $col2);


    # Write the packed data
    $self->_append( $header,
                    $data,
                    $unknown1,
                    $options,
                    $url_len,
                    $url);

    return $error;
}


###############################################################################
#
# _write_url_external($row1, $col1, $row2, $col2, $url, $string, $format)
#
# Write links to external directory names such as 'c:\foo.xls',
# c:\foo.xls#Sheet1!A1', '../../foo.xls'. and '../../foo.xls#Sheet1!A1'.
#
# Note: Excel writes some relative links with the $dir_long string. We ignore
# these cases for the sake of simpler code.
#
# See also write_url() above for a general description and return values.
#
sub _write_url_external {

    my $self    = shift;

    # Network drives are different. We will handle them separately
    # MS/Novell network drives and shares start with \\
    return $self->_write_url_external_net(@_) if $_[4] =~ m[^external:\\\\];


    my $record      = 0x01B8;                       # Record identifier
    my $length      = 0x00000;                      # Bytes to follow

    my $row1        = $_[0];                        # Start row
    my $col1        = $_[1];                        # Start column
    my $row2        = $_[2];                        # End row
    my $col2        = $_[3];                        # End column
    my $url         = $_[4];                        # URL string
    my $str         = $_[5];                        # Alternative label
    my $xf          = $_[6] || $self->{_url_format};# The cell format


    # Strip URL type and change Unix dir separator to Dos style (if needed)
    #
    $url            =~ s[^external:][];
    $url            =~ s[/][\\]g;


    # Write the visible label but protect against url recursion in write().
    ($str = $url)         =~ s[\#][ - ] unless defined $str;
    $self->{_writing_url} = 1;
    my $error             = $self->write($row1, $col1, $str, $xf);
    $self->{_writing_url} = 0;
    return $error         if $error == -2;


    # Determine if the link is relative or absolute:
    # Absolute if link starts with DOS drive specifier like C:
    # Otherwise default to 0x00 for relative link.
    #
    my $absolute    = 0x00;
       $absolute    = 0x02  if $url =~ m/^[A-Za-z]:/;


    # Determine if the link contains a sheet reference and change some of the
    # parameters accordingly.
    # Split the dir name and sheet name (if it exists)
    #
    my ($dir_long , $sheet) = split /\#/, $url;
    my $link_type           = 0x01 | $absolute;
    my $sheet_len;

    if (defined $sheet) {
        $link_type |= 0x08;
        $sheet_len  = pack("V", length($sheet) + 0x01);
        $sheet      = join("\0", split('', $sheet));
        $sheet     .= "\0\0\0";
    }
    else {
        $sheet_len  = '';
        $sheet      = '';
    }


    # Pack the link type
    $link_type      = pack("V", $link_type);


    # Calculate the up-level dir count e.g. (..\..\..\ == 3)
    my $up_count    = 0;
    $up_count++       while $dir_long =~ s[^\.\.\\][];
    $up_count       = pack("v", $up_count);


    # Store the short dos dir name (null terminated)
    my $dir_short   = $dir_long . "\0";


    # Store the long dir name as a wchar string (non-null terminated)
    $dir_long       = join("\0", split('', $dir_long));
    $dir_long       = $dir_long . "\0";


    # Pack the lengths of the dir strings
    my $dir_short_len = pack("V", length $dir_short      );
    my $dir_long_len  = pack("V", length $dir_long       );
    my $stream_len    = pack("V", length($dir_long) + 0x06);


    # Pack the undocumented parts of the hyperlink stream
    my $unknown1 =pack("H*",'D0C9EA79F9BACE118C8200AA004BA90B02000000'       );
    my $unknown2 =pack("H*",'0303000000000000C000000000000046'               );
    my $unknown3 =pack("H*",'FFFFADDE000000000000000000000000000000000000000');
    my $unknown4 =pack("v",  0x03                                            );


    # Pack the main data stream
    my $data        = pack("vvvv", $row1, $row2, $col1, $col2) .
                      $unknown1     .
                      $link_type    .
                      $unknown2     .
                      $up_count     .
                      $dir_short_len.
                      $dir_short    .
                      $unknown3     .
                      $stream_len   .
                      $dir_long_len .
                      $unknown4     .
                      $dir_long     .
                      $sheet_len    .
                      $sheet        ;


    # Pack the header data
    $length         = length $data;
    my $header      = pack("vv",   $record, $length);


    # Write the packed data
    $self->_append( $header, $data);

    return $error;
}




###############################################################################
#
# _write_url_external_net($row1, $col1, $row2, $col2, $url, $string, $format)
#
# Write links to external MS/Novell network drives and shares such as
# '//NETWORK/share/foo.xls' and '//NETWORK/share/foo.xls#Sheet1!A1'.
#
# See also write_url() above for a general description and return values.
#
sub _write_url_external_net {

    my $self    = shift;

    my $record      = 0x01B8;                       # Record identifier
    my $length      = 0x00000;                      # Bytes to follow

    my $row1        = $_[0];                        # Start row
    my $col1        = $_[1];                        # Start column
    my $row2        = $_[2];                        # End row
    my $col2        = $_[3];                        # End column
    my $url         = $_[4];                        # URL string
    my $str         = $_[5];                        # Alternative label
    my $xf          = $_[6] || $self->{_url_format};# The cell format


    # Strip URL type and change Unix dir separator to Dos style (if needed)
    #
    $url            =~ s[^external:][];
    $url            =~ s[/][\\]g;


    # Write the visible label but protect against url recursion in write().
    ($str = $url)         =~ s[\#][ - ] unless defined $str;
    $self->{_writing_url} = 1;
    my $error             = $self->write($row1, $col1, $str, $xf);
    $self->{_writing_url} = 0;
    return $error         if $error == -2;


    # Determine if the link contains a sheet reference and change some of the
    # parameters accordingly.
    # Split the dir name and sheet name (if it exists)
    #
    my ($dir_long , $sheet) = split /\#/, $url;
    my $link_type           = 0x0103; # Always absolute
    my $sheet_len;

    if (defined $sheet) {
        $link_type |= 0x08;
        $sheet_len  = pack("V", length($sheet) + 0x01);
        $sheet      = join("\0", split('', $sheet));
        $sheet     .= "\0\0\0";
    }
    else {
        $sheet_len   = '';
        $sheet       = '';
    }

    # Pack the link type
    $link_type      = pack("V", $link_type);


    # Make the string null terminated
    $dir_long       = $dir_long . "\0";


    # Pack the lengths of the dir string
    my $dir_long_len  = pack("V", length $dir_long);


    # Store the long dir name as a wchar string (non-null terminated)
    $dir_long       = join("\0", split('', $dir_long));
    $dir_long       = $dir_long . "\0";


    # Pack the undocumented part of the hyperlink stream
    my $unknown1    = pack("H*",'D0C9EA79F9BACE118C8200AA004BA90B02000000');


    # Pack the main data stream
    my $data        = pack("vvvv", $row1, $row2, $col1, $col2) .
                      $unknown1     .
                      $link_type    .
                      $dir_long_len .
                      $dir_long     .
                      $sheet_len    .
                      $sheet        ;


    # Pack the header data
    $length         = length $data;
    my $header      = pack("vv",   $record, $length);


    # Write the packed data
    $self->_append( $header, $data);

    return $error;
}


###############################################################################
#
# write_date_time ($row, $col, $string, $format)
#
# Write a datetime string in ISO8601 "yyyy-mm-ddThh:mm:ss.ss" format as a
# number representing an Excel date. $format is optional.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : Invalid date_time, written as string
#
sub write_date_time {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    if (@_ < 3) { return -1 }                        # Check the number of args

    my $row       = $_[0];                           # Zero indexed row
    my $col       = $_[1];                           # Zero indexed column
    my $str       = $_[2];


    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);

    my $error     = 0;
    my $date_time = $self->convert_date_time($str);

    if (defined $date_time) {
        $error = $self->write_number($row, $col, $date_time, $_[3]);
    }
    else {
        # The date isn't valid so write it as a string.
        $self->write_string($row, $col, $str, $_[3]);
        $error = -3;
    }
    return $error;
}



###############################################################################
#
# convert_date_time($date_time_string)
#
# The function takes a date and time in ISO8601 "yyyy-mm-ddThh:mm:ss.ss" format
# and converts it to a decimal number representing a valid Excel date.
#
# Dates and times in Excel are represented by real numbers. The integer part of
# the number stores the number of days since the epoch and the fractional part
# stores the percentage of the day in seconds. The epoch can be either 1900 or
# 1904.
#
# Parameter: Date and time string in one of the following formats:
#               yyyy-mm-ddThh:mm:ss.ss  # Standard
#               yyyy-mm-ddT             # Date only
#                         Thh:mm:ss.ss  # Time only
#
# Returns:
#            A decimal number representing a valid Excel date, or
#            undef if the date is invalid.
#
sub convert_date_time {

    my $self      = shift;
    my $date_time = $_[0];

    my $days      = 0; # Number of days since epoch
    my $seconds   = 0; # Time expressed as fraction of 24h hours in seconds

    my ($year, $month, $day);
    my ($hour, $min, $sec);


    # Strip leading and trailing whitespace.
    $date_time =~ s/^\s+//;
    $date_time =~ s/\s+$//;

    # Check for invalid date char.
    return if     $date_time =~ /[^0-9T:\-\.Z]/;

    # Check for "T" after date or before time.
    return unless $date_time =~ /\dT|T\d/;

    # Strip trailing Z in ISO8601 date.
    $date_time =~ s/Z$//;


    # Split into date and time.
    my ($date, $time) = split /T/, $date_time;


    # We allow the time portion of the input DateTime to be optional.
    if ($time ne '') {
        # Match hh:mm:ss.sss+ where the seconds are optional
        if ($time =~ /^(\d\d):(\d\d)(:(\d\d(\.\d+)?))?/) {
            $hour   = $1;
            $min    = $2;
            $sec    = $4 || 0;
        }
        else {
            return undef; # Not a valid time format.
        }

        # Some boundary checks
        return if $hour >= 24;
        return if $min  >= 60;
        return if $sec  >= 60;

        # Excel expresses seconds as a fraction of the number in 24 hours.
        $seconds = ($hour *60*60 + $min *60 + $sec) / (24 *60 *60);
    }


    # We allow the date portion of the input DateTime to be optional.
    return $seconds if $date eq '';


    # Match date as yyyy-mm-dd.
    if ($date =~ /^(\d\d\d\d)-(\d\d)-(\d\d)$/) {
        $year   = $1;
        $month  = $2;
        $day    = $3;
    }
    else {
        return undef; # Not a valid date format.
    }

    # Set the epoch as 1900 or 1904. Defaults to 1900.
    my $date_1904 = $self->{_1904};


    # Special cases for Excel.
    if (not $date_1904) {
        return      $seconds if $date eq '1899-12-31'; # Excel 1900 epoch
        return      $seconds if $date eq '1900-01-00'; # Excel 1900 epoch
        return 60 + $seconds if $date eq '1900-02-29'; # Excel false leapday
    }


    # We calculate the date by calculating the number of days since the epoch
    # and adjust for the number of leap days. We calculate the number of leap
    # days by normalising the year in relation to the epoch. Thus the year 2000
    # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
    #
    my $epoch   = $date_1904 ? 1904 : 1900;
    my $offset  = $date_1904 ?    4 :    0;
    my $norm    = 300;
    my $range   = $year -$epoch;


    # Set month days and check for leap year.
    my @mdays   = (31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
    my $leap    = 0;
       $leap    = 1  if $year % 4 == 0 and $year % 100 or $year % 400 == 0;
    $mdays[1]   = 29 if $leap;


    # Some boundary checks
    return if $year  < $epoch or $year  > 9999;
    return if $month < 1      or $month > 12;
    return if $day   < 1      or $day   > $mdays[$month -1];

    # Accumulate the number of days since the epoch.
    $days  = $day;                              # Add days for current month
    $days += $mdays[$_] for 0 .. $month -2;     # Add days for past months
    $days += $range *365;                       # Add days for past years
    $days += int(($range)                /  4); # Add leapdays
    $days -= int(($range +$offset)       /100); # Subtract 100 year leapdays
    $days += int(($range +$offset +$norm)/400); # Add 400 year leapdays
    $days -= $leap;                             # Already counted above


    # Adjust for Excel erroneously treating 1900 as a leap year.
    $days++ if $date_1904 == 0 and $days > 59;

    return $days + $seconds;
}





###############################################################################
#
# set_row($row, $height, $XF, $hidden, $level)
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
    my $grbit       = 0x0000;               # Option flags
    my $ixfe;                               # XF index
    my $height      = $_[1];                # Format object
    my $format      = $_[2];                # Format object
    my $hidden      = $_[3] || 0;           # Hidden flag
    my $level       = $_[4] || 0;           # Outline level


    return unless defined $rw;  # Ensure at least $row is specified.


    # Check for a format object
    if (ref $format) {
        $ixfe = $format->get_xf_index();
    }
    else {
        $ixfe = 0x0F;
    }


    # Set the row height in units of 1/20 of a point. Note, some heights may
    # not be obtained exactly due to rounding in Excel.
    #
    if (defined $height) {
        $miyRw = $height *20;
    }
    else {
        $miyRw = 0xff; # The default row height
        $height = 0;
    }


    # Set the limits for the outline levels (0 <= x <= 7).
    $level = 0 if $level < 0;
    $level = 7 if $level > 7;

    $self->{_outline_row_level} = $level if $level >$self->{_outline_row_level};


    # Set the options flags. fUnsynced is used to show that the font and row
    # heights are not compatible. This is usually the case for WriteExcel.
    # The collapsed flag 0x10 doesn't seem to be used to indicate that a row
    # is collapsed. Instead it is used to indicate that the previous row is
    # collapsed. The zero height flag, 0x20, is used to collapse a row.
    #
    $grbit |= $level;
    $grbit |= 0x0020 if $hidden;
    $grbit |= 0x0040; # fUnsynced
    $grbit |= 0x0080 if $format;
    $grbit |= 0x0100;


    my $header   = pack("vv",       $record, $length);
    my $data     = pack("vvvvvvvv", $rw, $colMic, $colMac, $miyRw,
                                    $irwMac,$reserved, $grbit, $ixfe);

    $self->_append($header, $data);

    # Store the row sizes for use when calculating image vertices.
    # Also store the column formats.
    $self->{_row_sizes}->{$_[0]}   = $height;
    $self->{_row_formats}->{$_[0]} = $format if defined $format;
}


###############################################################################
#
# _check_dimensions($row, $col)
#
# Check that $row and $col are valid and store max and min values for use in
# DIMENSIONS record. See, _store_dimensions().
#
sub _check_dimensions {

    my $self    = shift;
    my $row     = $_[0];
    my $col     = $_[1];

    if ($row >= $self->{_xls_rowmax}) { return -2 }
    if ($col >= $self->{_xls_colmax}) { return -2 }

    $self->{_dim_changed} = 1;

    if ($row <  $self->{_dim_rowmin}) { $self->{_dim_rowmin} = $row }
    if ($row >  $self->{_dim_rowmax}) { $self->{_dim_rowmax} = $row }
    if ($col <  $self->{_dim_colmin}) { $self->{_dim_colmin} = $col }
    if ($col >  $self->{_dim_colmax}) { $self->{_dim_colmax} = $col }

    return 0;
}


###############################################################################
#
# _store_dimensions()
#
# Writes Excel DIMENSIONS to define the area in which there is data.
#
sub _store_dimensions {

    my $self      = shift;
    my $record    = 0x0200;         # Record identifier
    my $length    = 0x000E;         # Number of bytes to follow
    my $row_min;                    # First row
    my $row_max;                    # Last row plus 1
    my $col_min;                    # First column
    my $col_max;                    # Last column plus 1
    my $reserved  = 0x0000;         # Reserved by Excel


    # Set the data range if data has been written to the worksheet
    if ($self->{_dim_changed}) {
        $row_min = $self->{_dim_rowmin};
        $row_max = $self->{_dim_rowmax} +1;
        $col_min = $self->{_dim_colmin};
        $col_max = $self->{_dim_colmax} +1;
    }
    else {
        # Special case, not data was written
        $row_min = 0;
        $row_max = 0;
        $col_min = 0;
        $col_max = 256;

    }


    my $header    = pack("vv",    $record, $length);
    my $data      = pack("VVvvv", $row_min, $row_max,
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

    use integer;    # Avoid << shift bug in Perl 5.6.0 on HP-UX

    my $self           = shift;
    my $record         = 0x023E;     # Record identifier
    my $length         = 0x0012;     # Number of bytes to follow

    my $grbit          = 0x00B6;     # Option flags
    my $rwTop          = $self->{_first_row};   # Top visible row
    my $colLeft        = $self->{_first_col};   # Leftmost visible column
    my $rgbHdr         = 0x00000040;            # Row/col heading, grid color

    my $wScaleSLV      = 0x0000;                # Zoom in page break preview
    my $wScaleNormal   = 0x0000;                # Zoom in normal view
    my $reserved       = 0x00000000;


    # The options flags that comprise $grbit
    my $fDspFmla       = $self->{_display_formulas}; # 0 - bit
    my $fDspGrid       = $self->{_screen_gridlines}; # 1
    my $fDspRwCol      = $self->{_display_headers};  # 2
    my $fFrozen        = $self->{_frozen};           # 3
    my $fDspZeros      = $self->{_display_zeros};    # 4
    my $fDefaultHdr    = 1;                          # 5
    my $fArabic        = $self->{_display_arabic};   # 6
    my $fDspGuts       = $self->{_outline_on};       # 7
    my $fFrozenNoSplit = 0;                          # 0 - bit
    my $fSelected      = $self->{_selected};         # 1
    my $fPaged         = $self->{_active};           # 2
    my $fBreakPreview  = 0;                          # 3

    $grbit             = $fDspFmla;
    $grbit            |= $fDspGrid       << 1;
    $grbit            |= $fDspRwCol      << 2;
    $grbit            |= $fFrozen        << 3;
    $grbit            |= $fDspZeros      << 4;
    $grbit            |= $fDefaultHdr    << 5;
    $grbit            |= $fArabic        << 6;
    $grbit            |= $fDspGuts       << 7;
    $grbit            |= $fFrozenNoSplit << 8;
    $grbit            |= $fSelected      << 9;
    $grbit            |= $fPaged         << 10;
    $grbit            |= $fBreakPreview  << 11;

    my $header  = pack("vv",      $record, $length);
    my $data    = pack("vvvVvvV", $grbit, $rwTop, $colLeft, $rgbHdr,
                                  $wScaleSLV, $wScaleNormal, $reserved );

    $self->_append($header, $data);
}


###############################################################################
#
# _store_page_view()
#
# Set page view mode. Only applicable to Mac Excel.
#
sub _store_page_view {

    my $self    = shift;

    return unless $self->{_page_view};

    my $data    = pack "H*", 'C8081100C808000000000040000000000900000000';

    $self->_append($data);
}


###############################################################################
#
# _store_tab_color()
#
# Write the Tab Color BIFF record.
#
sub _store_tab_color {

    my $self    = shift;
    my $color   = $self->{_tab_color};

    return unless $color;

    my $record  = 0x0862;      # Record identifier
    my $length  = 0x0014;      # Number of bytes to follow

    my $zero    = 0x0000;
    my $unknown = 0x0014;

    my $header  = pack("vv", $record, $length);
    my $data    = pack("vvvvvvvvvv", $record, $zero, $zero, $zero, $zero,
                                     $zero, $unknown, $zero, $color, $zero);

    $self->_append($header, $data);
}


###############################################################################
#
# _store_defrow()
#
# Write BIFF record DEFROWHEIGHT.
#
sub _store_defrow {

    my $self     = shift;
    my $record   = 0x0225;      # Record identifier
    my $length   = 0x0004;      # Number of bytes to follow

    my $grbit    = 0x0000;      # Options.
    my $height   = 0x00FF;      # Default row height

    my $header   = pack("vv", $record, $length);
    my $data     = pack("vv", $grbit,  $height);

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_defcol()
#
# Write BIFF record DEFCOLWIDTH.
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
    my $width    = $_[2] || 8.43;   # Col width in user units, 8.43 is default
    my $coldx;                      # Col width in internal units
    my $pixels;                     # Col width in pixels

    # Excel rounds the column width to the nearest pixel. Therefore we first
    # convert to pixels and then to the internal units. The pixel to users-units
    # relationship is different for values less than 1.
    #
    if ($width < 1) {
        $pixels = int($width *12);
    }
    else {
        $pixels = int($width *7 ) +5;
    }

    $coldx = int($pixels *256/7);


    my $ixfe;                       # XF index
    my $grbit    = 0x0000;          # Option flags
    my $reserved = 0x00;            # Reserved
    my $format   = $_[3];           # Format object
    my $hidden   = $_[4] || 0;      # Hidden flag
    my $level    = $_[5] || 0;      # Outline level


    # Check for a format object
    if (ref $format) {
        $ixfe = $format->get_xf_index();
    }
    else {
        $ixfe = 0x0F;
    }


    # Set the limits for the outline levels (0 <= x <= 7).
    $level = 0 if $level < 0;
    $level = 7 if $level > 7;


    # Set the options flags.
    # The collapsed flag 0x10 doesn't seem to be used to indicate that a col
    # is collapsed. Instead it is used to indicate that the previous col is
    # collapsed. The zero height flag, 0x20, is used to collapse a col.
    #
    $grbit |= 0x0001 if $hidden;
    $grbit |= $level << 8;


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

    my $pnn      = $self->{_active_pane};   # Pane position
    my $rwAct    = $_[0];                   # Active row
    my $colAct   = $_[1];                   # Active column
    my $irefAct  = 0;                       # Active cell ref
    my $cref     = 1;                       # Number of refs

    my $rwFirst  = $_[0];                   # First row in reference
    my $colFirst = $_[1];                   # First col in reference
    my $rwLast   = $_[2] || $rwFirst;       # Last  row in reference
    my $colLast  = $_[3] || $colFirst;      # Last  col in reference

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
# _store_externcount($count)
#
# Write BIFF record EXTERNCOUNT to indicate the number of external sheet
# references in a worksheet.
#
# Excel only stores references to external sheets that are used in formulas.
# For simplicity we store references to all the sheets in the workbook
# regardless of whether they are used or not. This reduces the overall
# complexity and eliminates the need for a two way dialogue between the formula
# parser the worksheet objects.
#
sub _store_externcount {

    my $self     = shift;
    my $record   = 0x0016;          # Record identifier
    my $length   = 0x0002;          # Number of bytes to follow

    my $cxals    = $_[0];           # Number of external references

    my $header   = pack("vv", $record, $length);
    my $data     = pack("v",  $cxals);

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_externsheet($sheetname)
#
#
# Writes the Excel BIFF EXTERNSHEET record. These references are used by
# formulas. A formula references a sheet name via an index. Since we store a
# reference to all of the external worksheets the EXTERNSHEET index is the same
# as the worksheet index.
#
sub _store_externsheet {

    my $self      = shift;

    my $record    = 0x0017;         # Record identifier
    my $length;                     # Number of bytes to follow

    my $sheetname = $_[0];          # Worksheet name
    my $cch;                        # Length of sheet name
    my $rgch;                       # Filename encoding

    # References to the current sheet are encoded differently to references to
    # external sheets.
    #
    if ($self->{_name} eq $sheetname) {
        $sheetname = '';
        $length    = 0x02;  # The following 2 bytes
        $cch       = 1;     # The following byte
        $rgch      = 0x02;  # Self reference
    }
    else {
        $length    = 0x02 + length($_[0]);
        $cch       = length($sheetname);
        $rgch      = 0x03;  # Reference to a sheet in the current workbook
    }

    my $header     = pack("vv",  $record, $length);
    my $data       = pack("CC", $cch, $rgch);

    $self->_prepend($header, $data, $sheetname);
}


###############################################################################
#
# _store_panes()
#
#
# Writes the Excel BIFF PANE record.
# The panes can either be frozen or thawed (unfrozen).
# Frozen panes are specified in terms of a integer number of rows and columns.
# Thawed panes are specified in terms of Excel's units for rows and columns.
#
sub _store_panes {

    my $self    = shift;
    my $record  = 0x0041;       # Record identifier
    my $length  = 0x000A;       # Number of bytes to follow

    my $y       = $_[0] || 0;   # Vertical split position
    my $x       = $_[1] || 0;   # Horizontal split position
    my $rwTop   = $_[2];        # Top row visible
    my $colLeft = $_[3];        # Leftmost column visible
    my $pnnAct  = $_[4];        # Active pane


    # Code specific to frozen or thawed panes.
    if ($self->{_frozen}) {
        # Set default values for $rwTop and $colLeft
        $rwTop   = $y unless defined $rwTop;
        $colLeft = $x unless defined $colLeft;
    }
    else {
        # Set default values for $rwTop and $colLeft
        $rwTop   = 0  unless defined $rwTop;
        $colLeft = 0  unless defined $colLeft;

        # Convert Excel's row and column units to the internal units.
        # The default row height is 12.75
        # The default column width is 8.43
        # The following slope and intersection values were interpolated.
        #
        $y = 20*$y      + 255;
        $x = 113.879*$x + 390;
    }


    # Determine which pane should be active. There is also the undocumented
    # option to override this should it be necessary: may be removed later.
    #
    if (not defined $pnnAct) {
        $pnnAct = 0 if ($x != 0 && $y != 0); # Bottom right
        $pnnAct = 1 if ($x != 0 && $y == 0); # Top right
        $pnnAct = 2 if ($x == 0 && $y != 0); # Bottom left
        $pnnAct = 3 if ($x == 0 && $y == 0); # Top left
    }

    $self->{_active_pane} = $pnnAct; # Used in _store_selection

    my $header     = pack("vv",    $record, $length);
    my $data       = pack("vvvvv", $x, $y, $rwTop, $colLeft, $pnnAct);

    $self->_append($header, $data);
}


###############################################################################
#
# _store_setup()
#
# Store the page setup SETUP BIFF record.
#
sub _store_setup {

    use integer;    # Avoid << shift bug in Perl 5.6.0 on HP-UX

    my $self         = shift;
    my $record       = 0x00A1;                  # Record identifier
    my $length       = 0x0022;                  # Number of bytes to follow


    # TODO some of these props don't have accessors. Add then as requested.

    my $iPaperSize   = $self->{_paper_size};    # Paper size
    my $iScale       = $self->{_print_scale};   # Print scaling factor
    my $iPageStart   = $self->{_page_start};    # Starting page number
    my $iFitWidth    = $self->{_fit_width};     # Fit to number of pages wide
    my $iFitHeight   = $self->{_fit_height};    # Fit to number of pages high
    my $grbit        = 0x00;                    # Option flags
    my $iRes         = 0x0258;                  # Print resolution
    my $iVRes        = 0x0258;                  # Vertical print resolution
    my $numHdr       = $self->{_margin_header}; # Header Margin
    my $numFtr       = $self->{_margin_footer}; # Footer Margin
    my $iCopies      = 0x01;                    # Number of copies


    my $fLeftToRight = $self->{_page_order};    # Print over then down
    my $fLandscape   = $self->{_orientation};   # Page orientation
    my $fNoPls       = 0x0;                     # Setup not read from printer
    my $fNoColor     = $self->{_black_white};   # Print black and white
    my $fDraft       = $self->{_draft_quality}; # Print draft quality
    my $fNotes       = $self->{_print_comments};# Print notes
    my $fNoOrient    = 0x0;                     # Orientation not set
    my $fUsePage     = 0x0;                     # Use custom starting page
       $fUsePage     = 0x1 if $self->{_page_start} > 1;


    $grbit           = $fLeftToRight;
    $grbit          |= $fLandscape    << 1;
    $grbit          |= $fNoPls        << 2;
    $grbit          |= $fNoColor      << 3;
    $grbit          |= $fDraft        << 4;
    $grbit          |= $fNotes        << 5;
    $grbit          |= $fNoOrient     << 6;
    $grbit          |= $fUsePage      << 7;


    $numHdr = pack("d", $numHdr);
    $numFtr = pack("d", $numFtr);

    if ($self->{_byte_order}) {
        $numHdr = reverse $numHdr;
        $numFtr = reverse $numFtr;
    }

    my $header = pack("vv",         $record, $length);
    my $data1  = pack("vvvvvvvv",   $iPaperSize,
                                    $iScale,
                                    $iPageStart,
                                    $iFitWidth,
                                    $iFitHeight,
                                    $grbit,
                                    $iRes,
                                    $iVRes);
    my $data2  = $numHdr .$numFtr;
    my $data3  = pack("v", $iCopies);

    $self->_prepend($header, $data1, $data2, $data3);

}

###############################################################################
#
# _store_header()
#
# Store the header caption BIFF record.
#
sub _store_header {

    my $self        = shift;

    my $record      = 0x0014;                       # Record identifier
    my $length;                                     # Bytes to follow

    my $str         = $self->{_header};             # header string
    my $cch         = length($str);                 # Length of header string
    my $encoding    = $self->{_header_encoding};    # Character encoding


    # Character length is num of chars not num of bytes
    $cch           /= 2 if $encoding;

    # Change the UTF-16 name from BE to LE
    $str            = pack 'n*', unpack 'v*', $str if $encoding;

    $length         = 3 + length($str);

    my $header      = pack("vv",  $record, $length);
    my $data        = pack("vC",  $cch, $encoding);

    $self->_prepend($header, $data, $str);
}


###############################################################################
#
# _store_footer()
#
# Store the footer caption BIFF record.
#
sub _store_footer {

    my $self        = shift;

    my $record      = 0x0015;                       # Record identifier
    my $length;                                     # Bytes to follow

    my $str         = $self->{_footer};             # footer string
    my $cch         = length($str);                 # Length of footer string
    my $encoding    = $self->{_footer_encoding};    # Character encoding


    # Character length is num of chars not num of bytes
    $cch           /= 2 if $encoding;

    # Change the UTF-16 name from BE to LE
    $str            = pack 'n*', unpack 'v*', $str if $encoding;

    $length         = 3 + length($str);

    my $header      = pack("vv",  $record, $length);
    my $data        = pack("vC",  $cch, $encoding);

    $self->_prepend($header, $data, $str);
}


###############################################################################
#
# _store_hcenter()
#
# Store the horizontal centering HCENTER BIFF record.
#
sub _store_hcenter {

    my $self     = shift;

    my $record   = 0x0083;              # Record identifier
    my $length   = 0x0002;              # Bytes to follow

    my $fHCenter = $self->{_hcenter};   # Horizontal centering

    my $header    = pack("vv",  $record, $length);
    my $data      = pack("v",   $fHCenter);

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_vcenter()
#
# Store the vertical centering VCENTER BIFF record.
#
sub _store_vcenter {

    my $self     = shift;

    my $record   = 0x0084;              # Record identifier
    my $length   = 0x0002;              # Bytes to follow

    my $fVCenter = $self->{_vcenter};   # Horizontal centering

    my $header    = pack("vv",  $record, $length);
    my $data      = pack("v",   $fVCenter);

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_margin_left()
#
# Store the LEFTMARGIN BIFF record.
#
sub _store_margin_left {

    my $self    = shift;

    my $record  = 0x0026;                   # Record identifier
    my $length  = 0x0008;                   # Bytes to follow

    my $margin  = $self->{_margin_left};    # Margin in inches

    my $header    = pack("vv",  $record, $length);
    my $data      = pack("d",   $margin);

    if ($self->{_byte_order}) { $data = reverse $data }

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_margin_right()
#
# Store the RIGHTMARGIN BIFF record.
#
sub _store_margin_right {

    my $self    = shift;

    my $record  = 0x0027;                   # Record identifier
    my $length  = 0x0008;                   # Bytes to follow

    my $margin  = $self->{_margin_right};   # Margin in inches

    my $header    = pack("vv",  $record, $length);
    my $data      = pack("d",   $margin);

    if ($self->{_byte_order}) { $data = reverse $data }

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_margin_top()
#
# Store the TOPMARGIN BIFF record.
#
sub _store_margin_top {

    my $self    = shift;

    my $record  = 0x0028;                   # Record identifier
    my $length  = 0x0008;                   # Bytes to follow

    my $margin  = $self->{_margin_top};     # Margin in inches

    my $header    = pack("vv",  $record, $length);
    my $data      = pack("d",   $margin);

    if ($self->{_byte_order}) { $data = reverse $data }

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_margin_bottom()
#
# Store the BOTTOMMARGIN BIFF record.
#
sub _store_margin_bottom {

    my $self    = shift;

    my $record  = 0x0029;                   # Record identifier
    my $length  = 0x0008;                   # Bytes to follow

    my $margin  = $self->{_margin_bottom};  # Margin in inches

    my $header    = pack("vv",  $record, $length);
    my $data      = pack("d",   $margin);

    if ($self->{_byte_order}) { $data = reverse $data }

    $self->_prepend($header, $data);
}


###############################################################################
#
# merge_cells($first_row, $first_col, $last_row, $last_col)
#
# This is an Excel97/2000 method. It is required to perform more complicated
# merging than the normal align merge in Format.pm
#
sub merge_cells {

    my $self    = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    my $record  = 0x00E5;                   # Record identifier
    my $length  = 0x000A;                   # Bytes to follow

    my $cref     = 1;                       # Number of refs
    my $rwFirst  = $_[0];                   # First row in reference
    my $colFirst = $_[1];                   # First col in reference
    my $rwLast   = $_[2] || $rwFirst;       # Last  row in reference
    my $colLast  = $_[3] || $colFirst;      # Last  col in reference


    # Excel doesn't allow a single cell to be merged
    return if $rwFirst == $rwLast and $colFirst == $colLast;

    # Swap last row/col with first row/col as necessary
    ($rwFirst,  $rwLast ) = ($rwLast,  $rwFirst ) if $rwFirst  > $rwLast;
    ($colFirst, $colLast) = ($colLast, $colFirst) if $colFirst > $colLast;

    my $header   = pack("vv",       $record, $length);
    my $data     = pack("vvvvv",    $cref,
                                    $rwFirst, $rwLast,
                                    $colFirst, $colLast);

    $self->_append($header, $data);
}


###############################################################################
#
# merge_range($row1, $col1, $row2, $col2, $string, $format, $encoding)
#
# This is a wrapper to ensure correct use of the merge_cells method, i.e., write
# the first cell of the range, write the formatted blank cells in the range and
# then call the merge_cells record. Failing to do the steps in this order will
# cause Excel 97 to crash.
#
sub merge_range {

    my $self    = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }
    croak "Incorrect number of arguments" if @_ != 6 and @_ != 7;
    croak "Format argument is not a format object" unless ref $_[5];

    my $rwFirst  = $_[0];
    my $colFirst = $_[1];
    my $rwLast   = $_[2];
    my $colLast  = $_[3];
    my $string   = $_[4];
    my $format   = $_[5];
    my $encoding = $_[6] ? 1 : 0;


    # Temp code to prevent merged formats in non-merged cells.
    my $error = "Error: refer to merge_range() in the documentation. " .
                "Can't use previously non-merged format in merged cells";

    croak $error if $format->{_used_merge} == -1;
    $format->{_used_merge} = 0; # Until the end of this function.


    # Set the merge_range property of the format object. For BIFF8+.
    $format->set_merge_range();

    # Excel doesn't allow a single cell to be merged
    croak "Can't merge single cell" if $rwFirst  == $rwLast and
                                       $colFirst == $colLast;

    # Swap last row/col with first row/col as necessary
    ($rwFirst,  $rwLast ) = ($rwLast,  $rwFirst ) if $rwFirst  > $rwLast;
    ($colFirst, $colLast) = ($colLast, $colFirst) if $colFirst > $colLast;

    # Write the first cell
    if ($encoding) {
        $self->write_unicode($rwFirst, $colFirst, $string, $format);
    }
    else {
        $self->write        ($rwFirst, $colFirst, $string, $format);
    }

    # Pad out the rest of the area with formatted blank cells.
    for my $row ($rwFirst .. $rwLast) {
        for my $col ($colFirst .. $colLast) {
            next if $row == $rwFirst and $col == $colFirst;
            $self->write_blank($row, $col, $format);
        }
    }

    $self->merge_cells($rwFirst, $colFirst, $rwLast, $colLast);

    # Temp code to prevent merged formats in non-merged cells.
    $format->{_used_merge} = 1;

}


###############################################################################
#
# _store_print_headers()
#
# Write the PRINTHEADERS BIFF record.
#
sub _store_print_headers {

    my $self        = shift;

    my $record      = 0x002a;                   # Record identifier
    my $length      = 0x0002;                   # Bytes to follow

    my $fPrintRwCol = $self->{_print_headers};  # Boolean flag

    my $header      = pack("vv",  $record, $length);
    my $data        = pack("v",   $fPrintRwCol);

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_print_gridlines()
#
# Write the PRINTGRIDLINES BIFF record. Must be used in conjunction with the
# GRIDSET record.
#
sub _store_print_gridlines {

    my $self        = shift;

    my $record      = 0x002b;                    # Record identifier
    my $length      = 0x0002;                    # Bytes to follow

    my $fPrintGrid  = $self->{_print_gridlines}; # Boolean flag

    my $header      = pack("vv",  $record, $length);
    my $data        = pack("v",   $fPrintGrid);

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_gridset()
#
# Write the GRIDSET BIFF record. Must be used in conjunction with the
# PRINTGRIDLINES record.
#
sub _store_gridset {

    my $self        = shift;

    my $record      = 0x0082;                        # Record identifier
    my $length      = 0x0002;                        # Bytes to follow

    my $fGridSet    = not $self->{_print_gridlines}; # Boolean flag

    my $header      = pack("vv",  $record, $length);
    my $data        = pack("v",   $fGridSet);

    $self->_prepend($header, $data);

}


###############################################################################
#
# _store_guts()
#
# Write the GUTS BIFF record. This is used to configure the gutter margins
# where Excel outline symbols are displayed. The visibility of the gutters is
# controlled by a flag in WSBOOL. See also _store_wsbool().
#
# We are all in the gutter but some of us are looking at the stars.
#
sub _store_guts {

    my $self        = shift;

    my $record      = 0x0080;   # Record identifier
    my $length      = 0x0008;   # Bytes to follow

    my $dxRwGut     = 0x0000;   # Size of row gutter
    my $dxColGut    = 0x0000;   # Size of col gutter

    my $row_level   = $self->{_outline_row_level};
    my $col_level   = 0;


    # Calculate the maximum column outline level. The equivalent calculation
    # for the row outline level is carried out in set_row().
    #
    foreach my $colinfo (@{$self->{_colinfo}}) {
        # Skip cols without outline level info.
        next if @{$colinfo} < 6;
        $col_level = @{$colinfo}[5] if @{$colinfo}[5] > $col_level;
    }


    # Set the limits for the outline levels (0 <= x <= 7).
    $col_level = 0 if $col_level < 0;
    $col_level = 7 if $col_level > 7;


    # The displayed level is one greater than the max outline levels
    $row_level++ if $row_level > 0;
    $col_level++ if $col_level > 0;

    my $header      = pack("vv",   $record, $length);
    my $data        = pack("vvvv", $dxRwGut, $dxColGut, $row_level, $col_level);

    $self->_prepend($header, $data);

}


###############################################################################
#
# _store_wsbool()
#
# Write the WSBOOL BIFF record, mainly for fit-to-page. Used in conjunction
# with the SETUP record.
#
sub _store_wsbool {

    my $self        = shift;

    my $record      = 0x0081;   # Record identifier
    my $length      = 0x0002;   # Bytes to follow

    my $grbit       = 0x0000;   # Option flags

    # Set the option flags
    $grbit |= 0x0001;                            # Auto page breaks visible
    $grbit |= 0x0020 if $self->{_outline_style}; # Auto outline styles
    $grbit |= 0x0040 if $self->{_outline_below}; # Outline summary below
    $grbit |= 0x0080 if $self->{_outline_right}; # Outline summary right
    $grbit |= 0x0100 if $self->{_fit_page};      # Page setup fit to page
    $grbit |= 0x0400 if $self->{_outline_on};    # Outline symbols displayed


    my $header      = pack("vv",  $record, $length);
    my $data        = pack("v",   $grbit);

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_hbreak()
#
# Write the HORIZONTALPAGEBREAKS BIFF record.
#
sub _store_hbreak {

    my $self    = shift;

    # Return if the user hasn't specified pagebreaks
    return unless @{$self->{_hbreaks}};

    # Sort and filter array of page breaks
    my @breaks  = $self->_sort_pagebreaks(@{$self->{_hbreaks}});

    my $record  = 0x001b;               # Record identifier
    my $cbrk    = scalar @breaks;       # Number of page breaks
    my $length  = 2 + 6*$cbrk;          # Bytes to follow


    my $header  = pack("vv",  $record, $length);
    my $data    = pack("v",   $cbrk);

    # Append each page break
    foreach my $break (@breaks) {
        $data .= pack("vvv", $break, 0x0000, 0x00ff);
    }

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_vbreak()
#
# Write the VERTICALPAGEBREAKS BIFF record.
#
sub _store_vbreak {

    my $self    = shift;

    # Return if the user hasn't specified pagebreaks
    return unless @{$self->{_vbreaks}};

    # Sort and filter array of page breaks
    my @breaks  = $self->_sort_pagebreaks(@{$self->{_vbreaks}});

    my $record  = 0x001a;               # Record identifier
    my $cbrk    = scalar @breaks;       # Number of page breaks
    my $length  = 2 + 6*$cbrk;          # Bytes to follow


    my $header  = pack("vv",  $record, $length);
    my $data    = pack("v",   $cbrk);

    # Append each page break
    foreach my $break (@breaks) {
        $data .= pack("vvv", $break, 0x0000, 0xffff);
    }

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_protect()
#
# Set the Biff PROTECT record to indicate that the worksheet is protected.
#
sub _store_protect {

    my $self        = shift;

    # Exit unless sheet protection has been specified
    return unless $self->{_protect};

    my $record      = 0x0012;               # Record identifier
    my $length      = 0x0002;               # Bytes to follow

    my $fLock       = $self->{_protect};    # Worksheet is protected

    my $header      = pack("vv", $record, $length);
    my $data        = pack("v",  $fLock);

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_obj_protect()
#
# Set the Biff OBJPROTECT record to indicate that objects are protected.
#
sub _store_obj_protect {

    my $self        = shift;

    # Exit unless sheet protection has been specified
    return unless $self->{_protect};

    my $record      = 0x0063;               # Record identifier
    my $length      = 0x0002;               # Bytes to follow

    my $fLock       = $self->{_protect};    # Worksheet is protected

    my $header      = pack("vv", $record, $length);
    my $data        = pack("v",  $fLock);

    $self->_prepend($header, $data);
}


###############################################################################
#
# _store_password()
#
# Write the worksheet PASSWORD record.
#
sub _store_password {

    my $self        = shift;

    # Exit unless sheet protection and password have been specified
    return unless $self->{_protect} and defined $self->{_password};

    my $record      = 0x0013;               # Record identifier
    my $length      = 0x0002;               # Bytes to follow

    my $wPassword   = $self->{_password};   # Encoded password

    my $header      = pack("vv", $record, $length);
    my $data        = pack("v",  $wPassword);

    $self->_prepend($header, $data);
}


###############################################################################
#
# insert_bitmap($row, $col, $filename, $x, $y, $scale_x, $scale_y)
#
# Insert a 24bit bitmap image in a worksheet. The main record required is
# IMDATA but it must be proceeded by a OBJ record to define its position.
#
sub insert_bitmap {

    my $self        = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    my $row         = $_[0];
    my $col         = $_[1];
    my $bitmap      = $_[2];
    my $x           = $_[3] || 0;
    my $y           = $_[4] || 0;
    my $scale_x     = $_[5] || 1;
    my $scale_y     = $_[6] || 1;

    my ($width, $height, $size, $data) = $self->_process_bitmap($bitmap);

    # Scale the frame of the image.
    $width  *= $scale_x;
    $height *= $scale_y;

    # Calculate the vertices of the image and write the OBJ record
    my @vertices = $self->_position_object($col, $row, $x, $y, $width, $height);
    $self->_store_obj_picture(@vertices );

    # Write the IMDATA record to store the bitmap data
    my $record      = 0x007f;
    my $length      = 8 + $size;
    my $cf          = 0x09;
    my $env         = 0x01;
    my $lcb         = $size;

    my $header      = pack("vvvvV", $record, $length, $cf, $env, $lcb);

    $self->_append($header, $data);
}


###############################################################################
#
#  _position_object()
#
# Calculate the vertices that define the position of a graphical object within
# the worksheet.
#
#         +------------+------------+
#         |     A      |      B     |
#   +-----+------------+------------+
#   |     |(x1,y1)     |            |
#   |  1  |(A1)._______|______      |
#   |     |    |              |     |
#   |     |    |              |     |
#   +-----+----|    BITMAP    |-----+
#   |     |    |              |     |
#   |  2  |    |______________.     |
#   |     |            |        (B2)|
#   |     |            |     (x2,y2)|
#   +---- +------------+------------+
#
# Example of a bitmap that covers some of the area from cell A1 to cell B2.
#
# Based on the width and height of the bitmap we need to calculate 8 vars:
#     $col_start, $row_start, $col_end, $row_end, $x1, $y1, $x2, $y2.
# The width and height of the cells are also variable and have to be taken into
# account.
# The values of $col_start and $row_start are passed in from the calling
# function. The values of $col_end and $row_end are calculated by subtracting
# the width and height of the bitmap from the width and height of the
# underlying cells.
# The vertices are expressed as a percentage of the underlying cell width as
# follows (rhs values are in pixels):
#
#       x1 = X / W *1024
#       y1 = Y / H *256
#       x2 = (X-1) / W *1024
#       y2 = (Y-1) / H *256
#
#       Where:  X is distance from the left side of the underlying cell
#               Y is distance from the top of the underlying cell
#               W is the width of the cell
#               H is the height of the cell
#
# Note: the SDK incorrectly states that the height should be expressed as a
# percentage of 1024.
#
sub _position_object {

    my $self = shift;

    my $col_start;  # Col containing upper left corner of object
    my $x1;         # Distance to left side of object

    my $row_start;  # Row containing top left corner of object
    my $y1;         # Distance to top of object

    my $col_end;    # Col containing lower right corner of object
    my $x2;         # Distance to right side of object

    my $row_end;    # Row containing bottom right corner of object
    my $y2;         # Distance to bottom of object

    my $width;      # Width of image frame
    my $height;     # Height of image frame

    ($col_start, $row_start, $x1, $y1, $width, $height) = @_;


    # Adjust start column for offsets that are greater than the col width
    while ($x1 >= $self->_size_col($col_start)) {
        $x1 -= $self->_size_col($col_start);
        $col_start++;
    }

    # Adjust start row for offsets that are greater than the row height
    while ($y1 >= $self->_size_row($row_start)) {
        $y1 -= $self->_size_row($row_start);
        $row_start++;
    }


    # Initialise end cell to the same as the start cell
    $col_end    = $col_start;
    $row_end    = $row_start;

    $width      = $width  + $x1 -1;
    $height     = $height + $y1 -1;


    # Subtract the underlying cell widths to find the end cell of the image
    while ($width >= $self->_size_col($col_end)) {
        $width -= $self->_size_col($col_end);
        $col_end++;
    }

    # Subtract the underlying cell heights to find the end cell of the image
    while ($height >= $self->_size_row($row_end)) {
        $height -= $self->_size_row($row_end);
        $row_end++;
    }

    # Bitmap isn't allowed to start or finish in a hidden cell, i.e. a cell
    # with zero eight or width.
    #
    return if $self->_size_col($col_start) == 0;
    return if $self->_size_col($col_end)   == 0;
    return if $self->_size_row($row_start) == 0;
    return if $self->_size_row($row_end)   == 0;

    # Convert the pixel values to the percentage value expected by Excel
    $x1 = $x1     / $self->_size_col($col_start)   * 1024;
    $y1 = $y1     / $self->_size_row($row_start)   *  256;
    $x2 = $width  / $self->_size_col($col_end)     * 1024;
    $y2 = $height / $self->_size_row($row_end)     *  256;

    # Simulate ceil() without calling POSIX::ceil().
    $x1 = int($x1 +0.5);
    $y1 = int($y1 +0.5);
    $x2 = int($x2 +0.5);
    $y2 = int($y2 +0.5);

    return( $col_start, $x1,
            $row_start, $y1,
            $col_end,   $x2,
            $row_end,   $y2
          );
}


###############################################################################
#
# _size_col($col)
#
# Convert the width of a cell from user's units to pixels. Excel rounds the
# column width to the nearest pixel. If the width hasn't been set by the user
# we use the default value. If the column is hidden we use a value of zero.
#
sub _size_col {

    my $self = shift;
    my $col  = $_[0];

    # Look up the cell value to see if it has been changed
    if (exists $self->{_col_sizes}->{$col}) {
        my $width = $self->{_col_sizes}->{$col};

        # The relationship is different for user units less than 1.
        if ($width < 1) {
            return int($width *12);
        }
        else {
            return int($width *7 ) +5;
        }
    }
    else {
        return 64;
    }
}


###############################################################################
#
# _size_row($row)
#
# Convert the height of a cell from user's units to pixels. By interpolation
# the relationship is: y = 4/3x. If the height hasn't been set by the user we
# use the default value. If the row is hidden we use a value of zero. (Not
# possible to hide row yet).
#
sub _size_row {

    my $self = shift;
    my $row  = $_[0];

    # Look up the cell value to see if it has been changed
    if (exists $self->{_row_sizes}->{$row}) {
        if ($self->{_row_sizes}->{$row} == 0) {
            return 0;
        }
        else {
            return int (4/3 * $self->{_row_sizes}->{$row});
        }
    }
    else {
        return 17;
    }
}


###############################################################################
#
# _store_obj_picture(   $col_start, $x1,
#                       $row_start, $y1,
#                       $col_end,   $x2,
#                       $row_end,   $y2 )
#
# Store the OBJ record that precedes an IMDATA record. This could be generalise
# to support other Excel objects.
#
sub _store_obj_picture {

    my $self        = shift;

    my $record      = 0x005d;   # Record identifier
    my $length      = 0x003c;   # Bytes to follow

    my $cObj        = 0x0001;   # Count of objects in file (set to 1)
    my $OT          = 0x0008;   # Object type. 8 = Picture
    my $id          = 0x0001;   # Object ID
    my $grbit       = 0x0614;   # Option flags

    my $colL        = $_[0];    # Col containing upper left corner of object
    my $dxL         = $_[1];    # Distance from left side of cell

    my $rwT         = $_[2];    # Row containing top left corner of object
    my $dyT         = $_[3];    # Distance from top of cell

    my $colR        = $_[4];    # Col containing lower right corner of object
    my $dxR         = $_[5];    # Distance from right of cell

    my $rwB         = $_[6];    # Row containing bottom right corner of object
    my $dyB         = $_[7];    # Distance from bottom of cell

    my $cbMacro     = 0x0000;   # Length of FMLA structure
    my $Reserved1   = 0x0000;   # Reserved
    my $Reserved2   = 0x0000;   # Reserved

    my $icvBack     = 0x09;     # Background colour
    my $icvFore     = 0x09;     # Foreground colour
    my $fls         = 0x00;     # Fill pattern
    my $fAuto       = 0x00;     # Automatic fill
    my $icv         = 0x08;     # Line colour
    my $lns         = 0xff;     # Line style
    my $lnw         = 0x01;     # Line weight
    my $fAutoB      = 0x00;     # Automatic border
    my $frs         = 0x0000;   # Frame style
    my $cf          = 0x0009;   # Image format, 9 = bitmap
    my $Reserved3   = 0x0000;   # Reserved
    my $cbPictFmla  = 0x0000;   # Length of FMLA structure
    my $Reserved4   = 0x0000;   # Reserved
    my $grbit2      = 0x0001;   # Option flags
    my $Reserved5   = 0x0000;   # Reserved


    my $header      = pack("vv", $record, $length);
    my $data        = pack("V",  $cObj);
       $data       .= pack("v",  $OT);
       $data       .= pack("v",  $id);
       $data       .= pack("v",  $grbit);
       $data       .= pack("v",  $colL);
       $data       .= pack("v",  $dxL);
       $data       .= pack("v",  $rwT);
       $data       .= pack("v",  $dyT);
       $data       .= pack("v",  $colR);
       $data       .= pack("v",  $dxR);
       $data       .= pack("v",  $rwB);
       $data       .= pack("v",  $dyB);
       $data       .= pack("v",  $cbMacro);
       $data       .= pack("V",  $Reserved1);
       $data       .= pack("v",  $Reserved2);
       $data       .= pack("C",  $icvBack);
       $data       .= pack("C",  $icvFore);
       $data       .= pack("C",  $fls);
       $data       .= pack("C",  $fAuto);
       $data       .= pack("C",  $icv);
       $data       .= pack("C",  $lns);
       $data       .= pack("C",  $lnw);
       $data       .= pack("C",  $fAutoB);
       $data       .= pack("v",  $frs);
       $data       .= pack("V",  $cf);
       $data       .= pack("v",  $Reserved3);
       $data       .= pack("v",  $cbPictFmla);
       $data       .= pack("v",  $Reserved4);
       $data       .= pack("v",  $grbit2);
       $data       .= pack("V",  $Reserved5);

    $self->_append($header, $data);
}


###############################################################################
#
# _process_bitmap()
#
# Convert a 24 bit bitmap into the modified internal format used by Windows.
# This is described in BITMAPCOREHEADER and BITMAPCOREINFO structures in the
# MSDN library.
#
sub _process_bitmap {

    my $self   = shift;
    my $bitmap = shift;

    # Open file and binmode the data in case the platform needs it.
    my $fh = FileHandle->new($bitmap);
    croak "Couldn't import $bitmap: $!" unless defined $fh;
    binmode $fh;


    # Slurp the file into a string.
    my $data = do {local $/; <$fh>};

    $fh->close;

    # Check that the file is big enough to be a bitmap.
    if (length $data <= 0x36) {
        croak "$bitmap doesn't contain enough data.";
    }


    # The first 2 bytes are used to identify the bitmap.
    if (unpack("A2", $data) ne "BM") {
        croak "$bitmap doesn't appear to be a valid bitmap image.";
    }


    # Remove bitmap data: ID.
    $data = substr $data, 2;


    # Read and remove the bitmap size. This is more reliable than reading
    # the data size at offset 0x22.
    #
    my $size   =  unpack "V", substr $data, 0, 4, "";
       $size  -=  0x36;   # Subtract size of bitmap header.
       $size  +=  0x0C;   # Add size of BIFF header.


    # Remove bitmap data: reserved, offset, header length.
    $data = substr $data, 12;


    # Read and remove the bitmap width and height. Verify the sizes.
    my ($width, $height) = unpack "V2", substr $data, 0, 8, "";

    if ($width > 0xFFFF) {
        croak "$bitmap: largest image width supported is 65k.";
    }

    if ($height > 0xFFFF) {
        croak "$bitmap: largest image height supported is 65k.";
    }

    # Read and remove the bitmap planes and bpp data. Verify them.
    my ($planes, $bitcount) = unpack "v2", substr $data, 0, 4, "";

    if ($bitcount != 24) {
        croak "$bitmap isn't a 24bit true color bitmap.";
    }

    if ($planes != 1) {
        croak "$bitmap: only 1 plane supported in bitmap image.";
    }


    # Read and remove the bitmap compression. Verify compression.
    my $compression = unpack "V", substr $data, 0, 4, "";

    if ($compression != 0) {
        croak "$bitmap: compression not supported in bitmap image.";
    }

    # Remove bitmap data: data size, hres, vres, colours, imp. colours.
    $data = substr $data, 20;

    # Add the BITMAPCOREHEADER data
    my $header  = pack("Vvvvv", 0x000c, $width, $height, 0x01, 0x18);
    $data       = $header . $data;

    return ($width, $height, $size, $data);
}


###############################################################################
#
# _store_zoom($zoom)
#
#
# Store the window zoom factor. This should be a reduced fraction but for
# simplicity we will store all fractions with a numerator of 100.
#
sub _store_zoom {

    my $self        = shift;

    # If scale is 100 we don't need to write a record
    return if $self->{_zoom} == 100;

    my $record      = 0x00A0;               # Record identifier
    my $length      = 0x0004;               # Bytes to follow

    my $header      = pack("vv", $record, $length   );
    my $data        = pack("vv", $self->{_zoom}, 100);

    $self->_append($header, $data);
}


###############################################################################
#
# write_unicode ($row, $col, $string, $format)
#
# Write a Unicode string to the specified row and column (zero indexed).
# $format is optional.
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long string truncated to 255 chars
#
sub write_unicode {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    if (@_ < 3) { return -1 }                        # Check the number of args

    my $record      = 0x00FD;                        # Record identifier
    my $length      = 0x000A;                        # Bytes to follow

    my $row         = $_[0];                         # Zero indexed row
    my $col         = $_[1];                         # Zero indexed column
    my $strlen      = length($_[2]);
    my $str         = $_[2];
    my $xf          = _XF($self, $row, $col, $_[3]); # The cell format
    my $encoding    = 0x1;
    my $str_error   = 0;

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);

    # Limit the utf16 string to the max number of chars (not bytes).
    if ($strlen > 32767* 2) {
        $str       = substr($str, 0, 32767*2);
        $str_error = -3;
    }


    my $num_bytes = length $str;
    my $num_chars = int($num_bytes / 2);


    # Check for a valid 2-byte char string.
    croak "Uneven number of bytes in Unicode string" if $num_bytes % 2;


    # Change from UTF16 big-endian to little endian
    $str = pack "v*", unpack "n*", $str;


    # Add the encoding and length header to the string.
    my $str_header  = pack("vC", $num_chars, $encoding);
    $str            = $str_header . $str;


    if (not exists ${$self->{_str_table}}->{$str}) {
        ${$self->{_str_table}}->{$str} = ${$self->{_str_unique}}++;
    }


    ${$self->{_str_total}}++;


    my $header = pack("vv",   $record, $length);
    my $data   = pack("vvvV", $row, $col, $xf, ${$self->{_str_table}}->{$str});

    $self->_append($header, $data);

    return $str_error;
}



###############################################################################
#
# write_unicode_le ($row, $col, $string, $format)
#
# Write a Unicode string to the specified row and column (zero indexed).
# $format is optional.
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long string truncated to 255 chars
#
# TODO Refactor. Too much code share with write_unicode().
#
sub write_unicode_le {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    if (@_ < 3) { return -1 }                        # Check the number of args

    my $record      = 0x00FD;                        # Record identifier
    my $length      = 0x000A;                        # Bytes to follow

    my $row         = $_[0];                         # Zero indexed row
    my $col         = $_[1];                         # Zero indexed column
    my $strlen      = length($_[2]);
    my $str         = $_[2];
    my $xf          = _XF($self, $row, $col, $_[3]); # The cell format
    my $encoding    = 0x1;
    my $str_error   = 0;

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);

    # Limit the utf16 string to the max number of chars (not bytes).
    if ($strlen > 32767* 2) {
        $str       = substr($str, 0, 32767*2);
        $str_error = -3;
    }


    my $num_bytes = length $str;
    my $num_chars = int($num_bytes / 2);


    # Check for a valid 2-byte char string.
    croak "Uneven number of bytes in Unicode string" if $num_bytes % 2;

    # Add the encoding and length header to the string.
    my $str_header  = pack("vC", $num_chars, $encoding);
       $str         = $str_header . $str;


    if (not exists ${$self->{_str_table}}->{$str}) {
        ${$self->{_str_table}}->{$str} = ${$self->{_str_unique}}++;
    }


    ${$self->{_str_total}}++;


    my $header = pack("vv",   $record, $length);
    my $data   = pack("vvvV", $row, $col, $xf, ${$self->{_str_table}}->{$str});

    $self->_append($header, $data);

    return $str_error;
}



#
# Methods related to comments and MSO objects.
#

###############################################################################
#
# _prepare_comments()
#
# Turn the HoH that stores the comments into an array for easier handling.
#
sub _prepare_comments {

    my $self    = shift;

    my $count   = 0;
    my @comments;


    # We sort the comments by row and column but that isn't strictly required.
    #
    my @rows = sort {$a <=> $b} keys %{$self->{_comments}};

    for my $row (@rows) {
        my @cols = sort {$a <=> $b} keys %{$self->{_comments}->{$row}};

        for my $col (@cols) {
            push @comments, $self->{_comments}->{$row}->{$col};
            $count++;
        }
    }

    $self->{_comments}       = {};
    $self->{_comments_array} = [@comments];

    return $count;
}


###############################################################################
#
# _store_comments()
#
# Store the collections of records that make up cell comments.
#
sub _store_comments {

    my $self            = shift;

    my $record          = 0x00EC;           # Record identifier
    my $length          = 0x0000;           # Bytes to follow

    my @ids             = @{$self->{_comments_ids}};
    my @comments        = @{$self->{_comments_array}};
    my $num_comments    = scalar @comments;
    my $spid            = shift @ids;

    return unless $num_comments;


    for my $i (0 .. $num_comments-1) {

        my $row         =   $comments[$i]->[0];
        my $col         =   $comments[$i]->[1];
        my $str         =   $comments[$i]->[2];
        my $encoding    =   $comments[$i]->[3];
        my $visible     =   $comments[$i]->[6];
        my $color       =   $comments[$i]->[7];
        my @vertices    = @{$comments[$i]->[8]};
        my @anchor_data = ($row, $col, @vertices);
        my $str_len     = length $str;
           $str_len    /= 2 if $encoding; # Num of chars not bytes.
        my $formats     = [[0, 5], [$str_len, 0]];


        if ($i == 0) {
            # Write the parent MSODRAWIING record.
            my $dg_length   = 200 + 128*($num_comments -1);
            my $spgr_length = 176 + 128*($num_comments -1);

            my $data        = $self->_store_mso_dg_container($dg_length);
               $data       .= $self->_store_mso_dg(@ids);
               $data       .= $self->_store_mso_spgr_container($spgr_length);
               $data       .= $self->_store_mso_sp_container(40);
               $data       .= $self->_store_mso_spgr();
               $data       .= $self->_store_mso_sp(0x0, $spid++, 0x0005);
               $data       .= $self->_store_mso_sp_container(120);
               $data       .= $self->_store_mso_sp(202, $spid++, 0x0A00);
               $data       .= $self->_store_mso_opt(0x80, $visible, $color);
               $data       .= $self->_store_mso_client_anchor(@anchor_data);
               $data       .= $self->_store_mso_client_data();

            $length         = length $data;
            my $header      = pack("vv", $record, $length);
            $self->_append($header, $data);

        }
        else {
            # Write the child MSODRAWIING record.
            my $data        = $self->_store_mso_sp_container(120);
               $data       .= $self->_store_mso_sp(202, $spid++, 0x0A00);
               $data       .= $self->_store_mso_opt(0x80, $visible,$color);
               $data       .= $self->_store_mso_client_anchor(@anchor_data);
               $data       .= $self->_store_mso_client_data();

            $length         = length $data;
            my $header      = pack("vv", $record, $length);
            $self->_append($header, $data);


        }

        $self->_store_obj_comment($i+1);
        $self->_store_mso_drawing_text_box();
        $self->_store_txo($str_len);
        $self->_store_txo_continue_1($str, $encoding);
        $self->_store_txo_continue_2($formats);
    }


    # Write the NOTE records after MSODRAWIING records.
    for my $i (0 .. $num_comments-1) {

        my $row         = $comments[$i]->[0];
        my $col         = $comments[$i]->[1];
        my $author      = $comments[$i]->[4];
        my $author_enc  = $comments[$i]->[5];
        my $visible     = $comments[$i]->[6];

        $self->_store_note($row, $col, $i+1, $author, $author_enc, $visible);
    }
}


###############################################################################
#
# _store_mso_dg_container()
#
# Write the Escher DgContainer record that is part of MSODRAWING.
#
sub _store_mso_dg_container {

    my $self        = shift;

    my $type        = 0xF002;
    my $version     = 15;
    my $instance    = 0;
    my $data        = '';
    my $length      = $_[0];


    return $self->_add_mso_generic($type, $version, $instance, $data, $length);
}


###############################################################################
#
# _store_mso_dg()
#
# Write the Escher Dg record that is part of MSODRAWING.
#
sub _store_mso_dg {

    my $self        = shift;

    my $type        = 0xF008;
    my $version     = 0;
    my $instance    = $_[0];
    my $data        = '';
    my $length      = 8;

    my $num_shapes  = $_[1];
    my $max_spid    = $_[2];

    $data           = pack "VV", $num_shapes, $max_spid;

    return $self->_add_mso_generic($type, $version, $instance, $data, $length);
}


###############################################################################
#
# _store_mso_spgr_container()
#
# Write the Escher SpgrContainer record that is part of MSODRAWING.
#
sub _store_mso_spgr_container {

    my $self        = shift;

    my $type        = 0xF003;
    my $version     = 15;
    my $instance    = 0;
    my $data        = '';
    my $length      = $_[0];


    return $self->_add_mso_generic($type, $version, $instance, $data, $length);
}


###############################################################################
#
# _store_mso_sp_container()
#
# Write the Escher SpContainer record that is part of MSODRAWING.
#
sub _store_mso_sp_container {

    my $self        = shift;

    my $type        = 0xF004;
    my $version     = 15;
    my $instance    = 0;
    my $data        = '';
    my $length      = $_[0];


    return $self->_add_mso_generic($type, $version, $instance, $data, $length);
}


###############################################################################
#
# _store_mso_spgr()
#
# Write the Escher Spgr record that is part of MSODRAWING.
#
sub _store_mso_spgr {

    my $self        = shift;

    my $type        = 0xF009;
    my $version     = 1;
    my $instance    = 0;
    my $data        = pack "VVVV", 0, 0, 0, 0;
    my $length      = 16;


    return $self->_add_mso_generic($type, $version, $instance, $data, $length);
}


###############################################################################
#
# _store_mso_sp()
#
# Write the Escher Sp record that is part of MSODRAWING.
#
sub _store_mso_sp {

    my $self        = shift;

    my $type        = 0xF00A;
    my $version     = 2;
    my $instance    = $_[0];
    my $data        = '';
    my $length      = 8;

    my $spid        = $_[1];
    my $options     = $_[2];

    $data           = pack "VV", $spid, $options;

    return $self->_add_mso_generic($type, $version, $instance, $data, $length);
}


###############################################################################
#
# _store_mso_opt()
#
# Write the Escher Opt record that is part of MSODRAWING.
#
sub _store_mso_opt {

    my $self        = shift;

    my $type        = 0xF00B;
    my $version     = 3;
    my $instance    = 9;
    my $data        = '';
    my $length      = 54;

    my $spid        = $_[0];
    my $visible     = $_[1];
    my $colour      = $_[2] || 0x50;


    # Use the visible flag if set by the user or else use the worksheet value.
    # Note that the value used is the opposite of _store_note().
    #
    if (defined $visible) {
        $visible = $visible                   ? 0x0000 : 0x0002;
    }
    else {
        $visible = $self->{_comments_visible} ? 0x0000 : 0x0002;
    }


    $data    = pack "V",  $spid;
    $data   .= pack "H*", '0000BF00080008005801000000008101' ;
    $data   .= pack "C",  $colour;
    $data   .= pack "H*", '000008830150000008BF011000110001' .
                          '02000000003F0203000300BF03';
    $data   .= pack "v",  $visible;
    $data   .= pack "H*", '0A00';


    return $self->_add_mso_generic($type, $version, $instance, $data, $length);
}


###############################################################################
#
# _store_mso_client_anchor()
#
# Write the Escher ClientAnchor record that is part of MSODRAWING.
#
sub _store_mso_client_anchor {

    my $self        = shift;

    my $type        = 0xF010;
    my $version     = 0;
    my $instance    = 0;
    my $data        = '';
    my $length      = 18;

    my $flag        = 3;
    my $row         = $_[0];
    my $col         = $_[1];


    my $col_start   = $_[2];    # Col containing upper left corner of object
    my $x1          = $_[3];    # Distance to left side of object

    my $row_start   = $_[4];    # Row containing top left corner of object
    my $y1          = $_[5];    # Distance to top of object

    my $col_end     = $_[6];    # Col containing lower right corner of object
    my $x2          = $_[7];    # Distance to right side of object

    my $row_end     = $_[8];    # Row containing bottom right corner of object
    my $y2          = $_[9];    # Distance to bottom of object

    my $width       = $_[10];   # Width of image frame
    my $height      = $_[11];   # Height of image frame


    $data   = pack "v9",    $flag,
                            $col_start, $x1,
                            $row_start, $y1,
                            $col_end,   $x2,
                            $row_end,   $y2;



    return $self->_add_mso_generic($type, $version, $instance, $data, $length);
}


###############################################################################
#
# _store_mso_client_data()
#
# Write the Escher ClientData record that is part of MSODRAWING.
#
sub _store_mso_client_data {

    my $self        = shift;

    my $type        = 0xF011;
    my $version     = 0;
    my $instance    = 0;
    my $data        = '';
    my $length      = 0;


    return $self->_add_mso_generic($type, $version, $instance, $data, $length);
}


###############################################################################
#
# _store_obj_comment()
#
# Write the OBJ record that is part of cell comments.
#
sub _store_obj_comment {

    my $self        = shift;

    my $record      = 0x005D;   # Record identifier
    my $length      = 0x0034;   # Bytes to follow

    my $obj_id      = $_[0];    # Object ID number.
    my $obj_type    = 0x0019;   # Object type (comment).
    my $data        = '';       # Record data.

    my $sub_record  = 0x0000;   # Sub-record identifier.
    my $sub_length  = 0x0000;   # Length of sub-record.
    my $sub_data    = '';       # Data of sub-record.
    my $options     = 0x4011;
    my $reserved    = 0x0000;

    # Add ftCmo (common object data) subobject
    $sub_record     = 0x0015;   # ftCmo
    $sub_length     = 0x0012;
    $sub_data       = pack "vvvVVV", $obj_type, $obj_id,   $options,
                                     $reserved, $reserved, $reserved;
    $data           = pack("vv",     $sub_record, $sub_length);
    $data          .= $sub_data;


    # Add ftNts (note structure) subobject
    $sub_record     = 0x000D;   # ftNts
    $sub_length     = 0x0016;
    $sub_data       = pack "VVVVVv", ($reserved) x 5;
    $data          .= pack("vv",     $sub_record, $sub_length);
    $data          .= $sub_data;


    # Add ftEnd (end of object) subobject
    $sub_record     = 0x0000;   # ftNts
    $sub_length     = 0x0000;
    $data          .= pack("vv",     $sub_record, $sub_length);


    # Pack the record.
    my $header  = pack("vv",        $record, $length);

    $self->_append($header, $data);

}


###############################################################################
#
# _store_mso_drawing_text_box()
#
# Write the MSODRAWING ClientTextbox record that is part of comments.
#
sub _store_mso_drawing_text_box {

    my $self        = shift;

    my $record      = 0x00EC;           # Record identifier
    my $length      = 0x0008;           # Bytes to follow


    my $data        = $self->_store_mso_client_text_box();
    my $header      = pack("vv", $record, $length);

    $self->_append($header, $data);
}


###############################################################################
#
# _store_mso_client_text_box()
#
# Write the Escher ClientTextbox record that is part of MSODRAWING.
#
sub _store_mso_client_text_box {

    my $self        = shift;

    my $type        = 0xF00D;
    my $version     = 0;
    my $instance    = 0;
    my $data        = '';
    my $length      = 0;


    return $self->_add_mso_generic($type, $version, $instance, $data, $length);
}


###############################################################################
#
# _store_txo()
#
# Write the worksheet TXO record that is part of cell comments.
#
sub _store_txo {

    my $self        = shift;

    my $record      = 0x01B6;               # Record identifier
    my $length      = 0x0012;               # Bytes to follow

    my $string_len  = $_[0];                # Length of the note text.
    my $format_len  = $_[1] || 16;          # Length of the format runs.
    my $rotation    = $_[2] || 0;           # Options
    my $grbit       = 0x0212;               # Options
    my $reserved    = 0x0000;               # Options

    # Pack the record.
    my $header  = pack("vv",        $record, $length);
    my $data    = pack("vvVvvvV",   $grbit, $rotation, $reserved, $reserved,
                                    $string_len, $format_len, $reserved);

    $self->_append($header, $data);

}


###############################################################################
#
# _store_txo_continue_1()
#
# Write the first CONTINUE record to follow the TXO record. It contains the
# text data.
#
sub _store_txo_continue_1 {

    my $self        = shift;

    my $record      = 0x003C;               # Record identifier
    my $length      = 0x0000;               # Bytes to follow

    my $string      = $_[0];                # Comment string.
    my $encoding    = $_[1] || 0;           # Encoding of the string.

    # Pack the record.
    my $data    = pack("C", $encoding) . $string;
       $length  = length $data;
    my $header  = pack("vv", $record, $length);

    $self->_append($header, $data);

}


###############################################################################
#
# _store_txo_continue_2()
#
# Write the second CONTINUE record to follow the TXO record. It contains the
# formatting information for the string.
#
sub _store_txo_continue_2 {

    my $self        = shift;

    my $record      = 0x003C;               # Record identifier
    my $length      = 0x0000;               # Bytes to follow
    my $formats     = $_[0];                # Formatting information


    # Pack the record.
    my $data = '';

    for my $a_ref (@$formats) {
        $data .= pack "vvV", $a_ref->[0], $a_ref->[1], 0x0;
    }

    $length     = length $data;
    my $header  = pack("vv", $record, $length);


    $self->_append($header, $data);

}


###############################################################################
#
# _store_note()
#
# Write the worksheet NOTE record that is part of cell comments.
#
sub _store_note {

    my $self        = shift;

    my $record      = 0x001C;               # Record identifier
    my $length      = 0x000C;               # Bytes to follow

    my $row         = $_[0];
    my $col         = $_[1];
    my $obj_id      = $_[2];
    my $author      = $_[3] || $self->{_comments_author};
    my $author_enc  = $_[4] || $self->{_comments_author_enc};
    my $visible     = $_[5];


    # Use the visible flag if set by the user or else use the worksheet value.
    # The flag is also set in _store_mso_opt() but with the opposite value.
    #
    if (defined $visible) {
        $visible = $visible                   ? 0x0002 : 0x0000;
    }
    else {
        $visible = $self->{_comments_visible} ? 0x0002 : 0x0000;
    }


    # Get the number of chars in the author string (not bytes).
    my $num_chars  = length $author;
       $num_chars /= 2 if $author_enc;


    # Null terminate the author string.
    $author .= "\0";


    # Pack the record.
    my $data    = pack("vvvvvC", $row, $col, $visible, $obj_id,
                                 $num_chars, $author_enc);

    $length     = length($data) + length($author);
    my $header  = pack("vv", $record, $length);

    $self->_append($header, $data, $author);
}


###############################################################################
#
# _comment_params()
#
# This method handles the additional optional parameters to write_comment() as
# well as calculating the comment object position and vertices.
#
sub _comment_params {

    my $self    = shift;

    my $row     = shift;
    my $col     = shift;
    my $string  = shift;

    my %params  = (
                    author          => '',
                    author_encoding => 0,
                    encoding        => 0,
                    color           => undef,
                    start_cell      => undef,
                    start_col       => undef,
                    start_row       => undef,
                    visible         => undef,
                    width           => 129,
                    height          => 75,
                    x_offset        => undef,
                    x_scale         => 1,
                    y_offset        => undef,
                    y_scale         => 1,
                  );


    # Overwrite the defaults with any user supplied values. Incorrect or
    # misspelled parameters are silently ignored.
    %params     = (%params, @_);


    # Ensure that a width and height have been set.
    $params{width}  = 129 if not $params{width};
    $params{height} = 75  if not $params{height};


    # Check that utf16 strings have an even number of bytes.
    if ($params{encoding}) {
        croak "Uneven number of bytes in comment string"
               if length($string) % 2;

        # Change from UTF-16BE to UTF-16LE
        $string = pack 'v*', unpack 'n*', $string;
    }

    if ($params{author_encoding}) {
        croak "Uneven number of bytes in author string"
                if length($params{author}) % 2;

        # Change from UTF-16BE to UTF-16LE
        $params{author} = pack 'v*', unpack 'n*', $params{author};
    }


    # Handle utf8 strings in perl 5.8.
    if ($] >= 5.008) {
        require Encode;

        if (Encode::is_utf8($string)) {
            $string = Encode::encode("UTF-16LE", $string);
            $params{encoding} = 1;
        }

        if (Encode::is_utf8($params{author})) {
            $params{author} = Encode::encode("UTF-16LE", $params{author});
            $params{author_encoding} = 1;
        }
    }


    # Limit the string to the max number of chars (not bytes).
    my $max_len  = 32767;
       $max_len *= 2 if $params{encoding};

    if (length($string) > $max_len) {
        $string       = substr($string, 0, $max_len);
    }


    # Set the comment background colour.
    my $color       = $params{color};
       $color       = &Spreadsheet::WriteExcel::Format::_get_color($color);
       $color       = 0x50 if $color == 0x7FFF; # Default color.
    $params{color}  = $color;


    # Convert a cell reference to a row and column.
    if (defined $params{start_cell}) {
        my ($row, $col)    = $self->_substitute_cellref($params{start_cell});
        $params{start_row} = $row;
        $params{start_col} = $col;
    }


    # Set the default start cell and offsets for the comment. These are
    # generally fixed in relation to the parent cell. However there are
    # some edge cases for cells at the, er, edges.
    #
    if (not defined $params{start_row}) {

        if    ($row == 0    ) {$params{start_row} = 0      }
        elsif ($row == 65533) {$params{start_row} = 65529  }
        elsif ($row == 65534) {$params{start_row} = 65530  }
        elsif ($row == 65535) {$params{start_row} = 65531  }
        else                  {$params{start_row} = $row -1}
    }

    if (not defined $params{y_offset}) {

        if    ($row == 0    ) {$params{y_offset}  = 2      }
        elsif ($row == 65533) {$params{y_offset}  = 4      }
        elsif ($row == 65534) {$params{y_offset}  = 4      }
        elsif ($row == 65535) {$params{y_offset}  = 2      }
        else                  {$params{y_offset}  = 7      }
    }

    if (not defined $params{start_col}) {

        if    ($col == 253  ) {$params{start_col} = 250    }
        elsif ($col == 254  ) {$params{start_col} = 251    }
        elsif ($col == 255  ) {$params{start_col} = 252    }
        else                  {$params{start_col} = $col +1}
    }

    if (not defined $params{x_offset}) {

        if    ($col == 253  ) {$params{x_offset}  = 49     }
        elsif ($col == 254  ) {$params{x_offset}  = 49     }
        elsif ($col == 255  ) {$params{x_offset}  = 49     }
        else                  {$params{x_offset}  = 15     }
    }


    # Scale the size of the comment box if required. We scale the width and
    # height using the relationship d2 =(d1 -1)*s +1, where d is dimension
    # and s is scale. This gives values that match Excel's behaviour.
    #
    if ($params{x_scale}) {
        $params{width}  = (($params{width}  -1) * $params{x_scale}) +1;
    }

    if ($params{y_scale}) {
        $params{height} = (($params{height} -1) * $params{y_scale}) +1;
    }


    # Calculate the positions of comment object.
    my @vertices = $self->_position_object( $params{start_col},
                                            $params{start_row},
                                            $params{x_offset},
                                            $params{y_offset},
                                            $params{width},
                                            $params{height}
                                          );

    return(
           $row,
           $col,
           $string,
           $params{encoding},
           $params{author},
           $params{author_encoding},
           $params{visible},
           $params{color},
           [@vertices]
          );
}




1;


__END__


=head1 NAME

Worksheet - A writer class for Excel Worksheets.

=head1 SYNOPSIS

See the documentation for Spreadsheet::WriteExcel

=head1 DESCRIPTION

This module is used in conjunction with Spreadsheet::WriteExcel.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMVI, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

