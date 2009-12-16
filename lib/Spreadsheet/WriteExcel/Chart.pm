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
    my $data  .= pack 'v', $width_basis;
       $data  .= pack 'v', $height_basis;
       $data  .= pack 'v', $height;
       $data  .= pack 'v', $scale_basis;
       $data  .= pack 'v', $index;

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
    my $data  .= pack 'V', $x_pos;
       $data  .= pack 'V', $y_pos;
       $data  .= pack 'V', $dx;
       $data  .= pack 'V', $dy;

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

� MM-MMIX, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

