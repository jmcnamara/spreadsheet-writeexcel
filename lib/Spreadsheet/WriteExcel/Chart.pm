package Spreadsheet::WriteExcel::Chart;

###############################################################################
#
# Chart - A writer class for Excel Charts.
#
#
# Used in conjunction with Spreadsheet::WriteExcel
#
# Copyright 2000-2009, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

use Exporter;
use strict;
use Carp;
use FileHandle;
use Spreadsheet::WriteExcel::Worksheet;


use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::WriteExcel::Worksheet);

$VERSION = '2.31';

###############################################################################
#
# new()
#
# Constructor. Creates a new Chart object from a Worksheet object
#
sub new {

    my $class = shift;
    my $self  = Spreadsheet::WriteExcel::Worksheet->new();

    $self->{_filename}     = $_[0];
    $self->{_name}         = $_[1];
    $self->{_index}        = $_[2];
    $self->{_encoding}     = $_[3];
    $self->{_activesheet}  = $_[4];
    $self->{_firstsheet}   = $_[5];
    $self->{_external_bin} = $_[6];
    $self->{_type}         = 0x0200;

    bless $self, $class;
    $self->_initialize();
    return $self;
}


###############################################################################
#
# _initialize()
#
# If we are handling the old-style external binary template then read the data
# into memory, otherwise use the SUPER _initialize().
#
#
sub _initialize {

    my $self = shift;

    if ( $self->{_external_bin} ) {
        my $filename   = $self->{_filename};
        my $filehandle = FileHandle->new($filename)
          or die "Couldn't open $filename in add_chart_ext(): $!.\n";

        binmode($filehandle);

        $self->{_filehandle}    = $filehandle;
        $self->{_datasize}      = -s $filehandle;
        $self->{_using_tmpfile} = 0;

        # Read the entire external chart binary into the the data buffer.
        # This will be retrieved by _get_data() when the chart is closed().
        read( $self->{_filehandle}, $self->{_data}, $self->{_datasize} );
    }
    else {
        $self->SUPER::_initialize();
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
}


###############################################################################
#
# BIFF Records.
#
###############################################################################


###############################################################################
#
# _store_3dbarshape()
#
# Write the 3DBARSHAPE chart BIFF record.
#
sub _store_3dbarshape {

    my $self = shift;

    my $record = 0x105F;    # Record identifier.
    my $length = 0x0002;    # Number of bytes to follow.
    my $riser  = 0x00;      # Shape of base.
    my $taper  = 0x00;      # Column taper type.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'C', $riser;
    $data .= pack 'C', $taper;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_ai()
#
# Write the AI chart BIFF record.
#
sub _store_ai {

    my $self = shift;

    my $record       = 0x1051;    # Record identifier.
    my $length       = 0x0008;    # Number of bytes to follow.
    my $id           = $_[0];     # Link index.
    my $type         = $_[1];     # Reference type.
    my $format_index = $_[2];     # Num format index.
    my $formula      = $_[3];     # Pre-parsed formula.
    my $grbit        = 0x0000;    # Option flags.

    my $formula_length = length $formula;

    $length += $formula_length;

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'C', $id;
    $data .= pack 'C', $type;
    $data .= pack 'v', $grbit;
    $data .= pack 'v', $format_index;
    $data .= pack 'v', $formula_length;
    $data .= $formula;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_areaformat()
#
# Write the AREAFORMAT chart BIFF record. Contains the patterns and colours
# of a chart area.
#
sub _store_areaformat {

    my $self = shift;

    my $record    = 0x100A;        # Record identifier.
    my $length    = 0x0010;        # Number of bytes to follow.
    my $rgbFore   = 0x00C0C0C0;    # Foreground RGB colour.
    my $rgbBack   = 0x00000000;    # Background RGB colour.
    my $pattern   = 0x0001;        # Pattern.
    my $grbit     = 0x0000;        # Option flags.
    my $indexFore = 0x0016;        # Index to Foreground colour.
    my $indexBack = 0x004F;        # Index to Background colour.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'V', $rgbFore;
    $data .= pack 'V', $rgbBack;
    $data .= pack 'v', $pattern;
    $data .= pack 'v', $grbit;
    $data .= pack 'v', $indexFore;
    $data .= pack 'v', $indexBack;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_axcext()
#
# Write the AXCEXT chart BIFF record.
#
sub _store_axcext {

    my $self = shift;

    my $record       = 0x1062;    # Record identifier.
    my $length       = 0x0012;    # Number of bytes to follow.
    my $catMin       = 0x0000;    # Minimum category on axis.
    my $catMax       = 0x0000;    # Maximum category on axis.
    my $catMajor     = 0x0001;    # Value of major unit.
    my $unitMajor    = 0x0000;    # Units of major unit.
    my $catMinor     = 0x0001;    # Value of minor unit.
    my $unitMinor    = 0x0000;    # Units of minor unit.
    my $unitBase     = 0x0000;    # Base unit of axis.
    my $catCrossDate = 0x0000;    # Crossing point.
    my $grbit        = 0x00EF;    # Option flags.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'v', $catMin;
    $data .= pack 'v', $catMax;
    $data .= pack 'v', $catMajor;
    $data .= pack 'v', $unitMajor;
    $data .= pack 'v', $catMinor;
    $data .= pack 'v', $unitMinor;
    $data .= pack 'v', $unitBase;
    $data .= pack 'v', $catCrossDate;
    $data .= pack 'v', $grbit;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_axesused()
#
# Write the AXESUSED chart BIFF record.
#
sub _store_axesused {

    my $self = shift;

    my $record   = 0x1046;    # Record identifier.
    my $length   = 0x0002;    # Number of bytes to follow.
    my $num_axes = $_[0];     # Number of axes used.

    my $header = pack 'vv', $record, $length;
    my $data = pack 'v', $num_axes;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_axis()
#
# Write the AXIS chart BIFF record tp define the axis type.
#
sub _store_axis {

    my $self = shift;

    my $record    = 0x101D;        # Record identifier.
    my $length    = 0x0012;        # Number of bytes to follow.
    my $type      = $_[0];         # Axis type.
    my $reserved1 = 0x00000000;    # Reserved.
    my $reserved2 = 0x00000000;    # Reserved.
    my $reserved3 = 0x00000000;    # Reserved.
    my $reserved4 = 0x00000000;    # Reserved.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'v', $type;
    $data .= pack 'V', $reserved1;
    $data .= pack 'V', $reserved2;
    $data .= pack 'V', $reserved3;
    $data .= pack 'V', $reserved4;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_axislineformat()
#
# Write the AXISLINEFORMAT chart BIFF record.
#
sub _store_axislineformat {

    my $self = shift;

    my $record      = 0x1021;    # Record identifier.
    my $length      = 0x0002;    # Number of bytes to follow.
    my $line_format = 0x0001;    # Axis line format.

    my $header = pack 'vv', $record, $length;
    my $data = pack 'v', $line_format;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_axisparent()
#
# Write the AXISPARENT chart BIFF record.
#
sub _store_axisparent {

    my $self = shift;

    my $record = 0x1041;        # Record identifier.
    my $length = 0x0012;        # Number of bytes to follow.
    my $iax    = $_[0];         # Axis index.
    my $x      = 0x000000A0;    # X-coord.
    my $y      = 0x00000099;    # Y-coord.
    my $dx     = 0x00000DB2;    # Length of x axis.
    my $dy     = 0x00000DE4;    # Length of y axis.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'v', $iax;
    $data .= pack 'V', $x;
    $data .= pack 'V', $y;
    $data .= pack 'V', $dx;
    $data .= pack 'V', $dy;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_bar()
#
# Write the BAR chart BIFF record. Defines a bar or column group.
#
sub _store_bar {

    my $self = shift;

    my $record    = 0x1017;    # Record identifier.
    my $length    = 0x0006;    # Number of bytes to follow.
    my $pcOverlap = 0x0000;    # Space between bars.
    my $pcGap     = 0x0096;    # Space between cats.
    my $grbit     = 0x0000;    # Option flags.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'v', $pcOverlap;
    $data .= pack 'v', $pcGap;
    $data .= pack 'v', $grbit;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_begin()
#
# Write the BEGIN chart BIFF record to indicate the start of a sub stream.
#
sub _store_begin {

    my $self = shift;

    my $record = 0x1033;    # Record identifier.
    my $length = 0x0000;    # Number of bytes to follow.

    my $header = pack 'vv', $record, $length;

    $self->_append($header);
}


###############################################################################
#
# _store_catserrange()
#
# Write the CATSERRANGE chart BIFF record.
#
sub _store_catserrange {

    my $self = shift;

    my $record   = 0x1020;    # Record identifier.
    my $length   = 0x0008;    # Number of bytes to follow.
    my $catCross = 0x0001;    # Value/category crossing.
    my $catLabel = 0x0001;    # Frequency of labels.
    my $catMark  = 0x0001;    # Frequency of ticks.
    my $grbit    = 0x0001;    # Option flags.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'v', $catCross;
    $data .= pack 'v', $catLabel;
    $data .= pack 'v', $catMark;
    $data .= pack 'v', $grbit;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_chart()
#
# Write the CHART BIFF record. This indicates the start of the chart sub-stream
# and contains dimensions of the chart on the display. Units are in 1/72 inch
# and are 2 byte integer with 2 byte fraction.
#
sub _store_chart {

    my $self = shift;

    my $record = 0x1002;        # Record identifier.
    my $length = 0x0010;        # Number of bytes to follow.
    my $x_pos  = 0x00000000;    # X pos of top left corner.
    my $y_pos  = 0x00000000;    # Y pos of top left corner.
    my $dx     = 0x02DD51E0;    # X size.
    my $dy     = 0x01C2B838;    # Y size.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'V', $x_pos;
    $data .= pack 'V', $y_pos;
    $data .= pack 'V', $dx;
    $data .= pack 'V', $dy;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_chartformat()
#
# Write the CHARTFORMAT chart BIFF record. The partent records for a formatting
# of a chart group.
#
sub _store_chartformat {

    my $self = shift;

    my $record    = 0x1014;        # Record identifier.
    my $length    = 0x0014;        # Number of bytes to follow.
    my $reserved1 = 0x00000000;    # Reserved.
    my $reserved2 = 0x00000000;    # Reserved.
    my $reserved3 = 0x00000000;    # Reserved.
    my $reserved4 = 0x00000000;    # Reserved.
    my $grbit     = 0x0000;        # Option flags.
    my $icrt      = 0x0000;        # Drawing order.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'V', $reserved1;
    $data .= pack 'V', $reserved2;
    $data .= pack 'V', $reserved3;
    $data .= pack 'V', $reserved4;
    $data .= pack 'v', $grbit;
    $data .= pack 'v', $icrt;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_charttext()
#
# Write the TEXT chart BIFF record.
#
sub _store_charttext {

    my $self = shift;

    my $record           = 0x1025;        # Record identifier.
    my $length           = 0x0020;        # Number of bytes to follow.
    my $horz_align       = 0x02;          # Horizontal alignment.
    my $vert_align       = 0x02;          # Vertical alignment.
    my $bg_mode          = 0x0001;        # Background display.
    my $text_color_rgb   = 0x00000000;    # Text RGB colour.
    my $text_x           = 0xFFFFFFEA;    # Text x-pos.
    my $text_y           = 0xFFFFFFDC;    # Text y-pos.
    my $text_dx          = 0x00000000;    # Width.
    my $text_dy          = 0x00000000;    # Height.
    my $grbit1           = 0x00B1;        # Options
    my $text_color_index = 0x004D;        # Auto Colour.
    my $grbit2           = 0x1020;        # Data label placement.
    my $rotation         = 0x0000;        # Text rotation.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'C', $horz_align;
    $data .= pack 'C', $vert_align;
    $data .= pack 'v', $bg_mode;
    $data .= pack 'V', $text_color_rgb;
    $data .= pack 'V', $text_x;
    $data .= pack 'V', $text_y;
    $data .= pack 'V', $text_dx;
    $data .= pack 'V', $text_dy;
    $data .= pack 'v', $grbit1;
    $data .= pack 'v', $text_color_index;
    $data .= pack 'v', $grbit2;
    $data .= pack 'v', $rotation;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_dataformat()
#
# Write the DATAFORMAT chart BIFF record. This record specifies the series
# that the subsequent sub stream refers to.
#
sub _store_dataformat {

    my $self = shift;

    my $record        = 0x1006;    # Record identifier.
    my $length        = 0x0008;    # Number of bytes to follow.
    my $point_number  = 0xFFFF;    # Point number.
    my $series_index  = $_[0];     # Series index.
    my $series_number = $_[0];     # Series number. (Same as index).
    my $grbit         = 0x0000;    # Format flags.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'v', $point_number;
    $data .= pack 'v', $series_index;
    $data .= pack 'v', $series_number;
    $data .= pack 'v', $grbit;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_defaulttext()
#
# Write the DEFAULTTEXT chart BIFF record. Identifier for subsequent TEXT
# record.
#
sub _store_defaulttext {

    my $self = shift;

    my $record = 0x1024;    # Record identifier.
    my $length = 0x0002;    # Number of bytes to follow.
    my $type   = 0x0002;    # Type.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'v', $type;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_end()
#
# Write the END chart BIFF record to indicate the end of a sub stream.
#
sub _store_end {

    my $self = shift;

    my $record = 0x1034;    # Record identifier.
    my $length = 0x0000;    # Number of bytes to follow.

    my $header = pack 'vv', $record, $length;

    $self->_append($header);
}


###############################################################################
#
# _store_fbi()
#
# Write the FBI chart BIFF record. Specifies the font information at the time
# it was applied to the chart.
#
sub _store_fbi {

    my $self = shift;

    my $record       = 0x1060;    # Record identifier.
    my $length       = 0x000A;    # Number of bytes to follow.
    my $index        = $_[0];     # Font index.
    my $height       = 0x00C8;    # Default font height in twips.
    my $width_basis  = 0x38B8;    # Width basis, in twips.
    my $height_basis = 0x22A1;    # Height basis, in twips.
    my $scale_basis  = 0x0000;    # Scale by chart area or plot area.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'v', $width_basis;
    $data .= pack 'v', $height_basis;
    $data .= pack 'v', $height;
    $data .= pack 'v', $scale_basis;
    $data .= pack 'v', $index;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_fontx()
#
# Write the FONTX chart BIFF record which contains the index of the FONT
# record in the Workbook.
#
sub _store_fontx {

    my $self = shift;

    my $record = 0x1026;    # Record identifier.
    my $length = 0x0002;    # Number of bytes to follow.
    my $index  = $_[0];     # Font index.

    my $header = pack 'vv', $record, $length;
    my $data = pack 'v', $index;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_frame()
#
# Write the FRAME chart BIFF record.
#
sub _store_frame {

    my $self = shift;

    my $record     = 0x1032;    # Record identifier.
    my $length     = 0x0004;    # Number of bytes to follow.
    my $frame_type = 0x0000;    # Frame type.
    my $grbit      = 0x0003;    # Option flags.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'v', $frame_type;
    $data .= pack 'v', $grbit;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_legend()
#
# Write the LEGEND chart BIFF record. The Marcus Horan method.
#
sub _store_legend {

    my $self = shift;

    my $record   = 0x1015;        # Record identifier.
    my $length   = 0x0014;        # Number of bytes to follow.
    my $x        = 0x00000E83;    # X-position.
    my $y        = 0x000006F9;    # Y-position.
    my $width    = 0x0000010B;    # Width.
    my $height   = 0x0000011C;    # Height.
    my $wType    = 0x03;          # Type.
    my $wSpacing = 0x01;          # Spacing.
    my $grbit    = 0x001F;        # Option flags.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'V', $x;
    $data .= pack 'V', $y;
    $data .= pack 'V', $width;
    $data .= pack 'V', $height;
    $data .= pack 'C', $wType;
    $data .= pack 'C', $wSpacing;
    $data .= pack 'v', $grbit;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_lineformat()
#
# Write the LINEFORMAT chart BIFF record.
#
sub _store_lineformat {

    my $self = shift;

    my $record = 0x1007;        # Record identifier.
    my $length = 0x000C;        # Number of bytes to follow.
    my $rgb    = 0x00000000;    # Line RGB colour.
    my $lns    = 0x0000;        # Line pattern.
    my $we     = 0xFFFF;        # Line weight.
    my $grbit  = 0x0009;        # Option flags.
    my $index  = 0x004D;        # Index to colour of line.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'V', $rgb;
    $data .= pack 'v', $lns;
    $data .= pack 'v', $we;
    $data .= pack 'v', $grbit;
    $data .= pack 'v', $index;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_plotarea()
#
# Write the PLOTAREA chart BIFF record. This indicates that the subsequent
# FRAME record belongs to a plot area.
#
sub _store_plotarea {

    my $self = shift;

    my $record = 0x1035;    # Record identifier.
    my $length = 0x0000;    # Number of bytes to follow.

    my $header = pack 'vv', $record, $length;

    $self->_append($header);
}


###############################################################################
#
# _store_pos()
#
# Write the POS chart BIFF record. Generally not required when using
# automatic positioning.
#
sub _store_pos {

    my $self = shift;

    my $record  = 0x104F;    # Record identifier.
    my $length  = 0x0014;    # Number of bytes to follow.
    my $mdTopLt = $_[0];     # Top left.
    my $mdBotRt = $_[1];     # Bottom right.
    my $x1      = $_[2];     # X coordinate.
    my $y1      = $_[3];     # Y coordinate.
    my $x2      = $_[4];     # Width.
    my $y2      = $_[5];     # Height.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'v', $mdTopLt;
    $data .= pack 'v', $mdBotRt;
    $data .= pack 'V', $x1;
    $data .= pack 'V', $y1;
    $data .= pack 'V', $x2;
    $data .= pack 'V', $y2;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_series()
#
# Write the SERIES chart BIFF record.
#
sub _store_series {

    my $self = shift;

    my $record         = 0x1003;    # Record identifier.
    my $length         = 0x000C;    # Number of bytes to follow.
    my $category_type  = $_[0];     # Type: category.
    my $value_type     = $_[1];     # Type: value.
    my $category_count = $_[2];     # Num of categories.
    my $value_count    = $_[3];     # Num of values.
    my $bubble_type    = $_[4];     # Type: bubble.
    my $bubble_count   = $_[5];     # Num of bubble values.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'v', $category_type;
    $data .= pack 'v', $value_type;
    $data .= pack 'v', $category_count;
    $data .= pack 'v', $value_count;
    $data .= pack 'v', $bubble_type;
    $data .= pack 'v', $bubble_count;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_sertocrt()
#
# Write the SERTOCRT chart BIFF record to indicate the chart group index.
#
sub _store_sertocrt {

    my $self = shift;

    my $record     = 0x1045;    # Record identifier.
    my $length     = 0x0002;    # Number of bytes to follow.
    my $chartgroup = 0x0000;    # Chart group index.

    my $header = pack 'vv', $record, $length;
    my $data = pack 'v', $chartgroup;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_shtprops()
#
# Write the SHTPROPS chart BIFF record.
#
sub _store_shtprops {

    my $self = shift;

    my $record      = 0x1044;    # Record identifier.
    my $length      = 0x0004;    # Number of bytes to follow.
    my $grbit       = 0x000E;    # Option flags.
    my $empty_cells = 0x0000;    # Empty cell handling.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'v', $grbit;
    $data .= pack 'v', $empty_cells;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_text()
#
# Write the TEXT chart BIFF record.
#
sub _store_text {

    my $self = shift;

    my $record   = 0x1025;        # Record identifier.
    my $length   = 0x0020;        # Number of bytes to follow.
    my $at       = 0x02;          # Horizontal alignment.
    my $vat      = 0x02;          # Vertical alignment.
    my $wBkgMode = 0x0001;        # Background display.
    my $rgbText  = 0x00000000;    # Text RGB colour.
    my $x        = 0xFFFFFFEA;    # Text x-pos.
    my $y        = 0xFFFFFFDC;    # Text y-pos.
    my $dx       = 0x00000000;    # Width.
    my $dy       = 0x00000000;    # Height.
    my $grbit    = 0x00B1;        # Option flags.
    my $icvText  = 0x004D;        # Auto Colour.
    my $grbit2   = 0x0000;        # Show legend.
    my $rotation = 0x0000;        # Show value.


    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'C', $at;
    $data .= pack 'C', $vat;
    $data .= pack 'v', $wBkgMode;
    $data .= pack 'V', $rgbText;
    $data .= pack 'V', $x;
    $data .= pack 'V', $y;
    $data .= pack 'V', $dx;
    $data .= pack 'V', $dy;
    $data .= pack 'v', $grbit;
    $data .= pack 'v', $icvText;
    $data .= pack 'v', $grbit2;
    $data .= pack 'v', $rotation;

    $self->_append( $header, $data );
}

###############################################################################
#
# _store_tick()
#
# Write the TICK chart BIFF record.
#
sub _store_tick {

    my $self = shift;

    my $record    = 0x101E;        # Record identifier.
    my $length    = 0x001E;        # Number of bytes to follow.
    my $tktMajor  = 0x02;          # Type of major tick mark.
    my $tktMinor  = 0x00;          # Type of minor tick mark.
    my $tlt       = 0x03;          # Tick label position.
    my $wBkgMode  = 0x01;          # Background mode.
    my $rgb       = 0x00000000;    # Tick-label RGB colour.
    my $reserved1 = 0x00000000;    # Reserved.
    my $reserved2 = 0x00000000;    # Reserved.
    my $reserved3 = 0x00000000;    # Reserved.
    my $reserved4 = 0x00000000;    # Reserved.
    my $grbit     = 0x0023;        # Option flags.
    my $index     = 0x004D;        # Colour index.
    my $reserved5 = 0x0000;        # Reserved.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'C', $tktMajor;
    $data .= pack 'C', $tktMinor;
    $data .= pack 'C', $tlt;
    $data .= pack 'C', $wBkgMode;
    $data .= pack 'V', $rgb;
    $data .= pack 'V', $reserved1;
    $data .= pack 'V', $reserved2;
    $data .= pack 'V', $reserved3;
    $data .= pack 'V', $reserved4;
    $data .= pack 'v', $grbit;
    $data .= pack 'v', $index;
    $data .= pack 'v', $reserved5;

    $self->_append( $header, $data );
}


###############################################################################
#
# _store_valuerange()
#
# Write the VALUERANGE chart BIFF record.
#
sub _store_valuerange {

    my $self = shift;

    my $record   = 0x101F;        # Record identifier.
    my $length   = 0x002A;        # Number of bytes to follow.
    my $numMin   = 0x00000000;    # Minimum value on axis.
    my $numMax   = 0x00000000;    # Maximum value on axis.
    my $numMajor = 0x00000000;    # Value of major increment.
    my $numMinor = 0x00000000;    # Value of minor increment.
    my $numCross = 0x00000000;    # Value where category axis crosses.
    my $grbit    = 0x011F;        # Format flags.

    # TODO. Reverse doubles when they are handled.

    my $header = pack 'vv', $record, $length;
    my $data = '';
    $data .= pack 'd', $numMin;
    $data .= pack 'd', $numMax;
    $data .= pack 'd', $numMajor;
    $data .= pack 'd', $numMinor;
    $data .= pack 'd', $numCross;
    $data .= pack 'v', $grbit;

    $self->_append( $header, $data );
}


1;


__END__


=head1 NAME

Chart - A writer class for Excel Charts.

=head1 SYNOPSIS

See the documentation for Spreadsheet::WriteExcel

=head1 DESCRIPTION

This module is used in conjunction with Spreadsheet::WriteExcel.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

Copyright MM-MMIX, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

