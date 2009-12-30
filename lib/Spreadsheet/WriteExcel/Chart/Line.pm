package Spreadsheet::WriteExcel::Chart::Line;

###############################################################################
#
# Line - A writer class for Excel Line charts.
#
# Used in conjunction with Spreadsheet::WriteExcel::Chart.
#
# See formatting note in Spreadsheet::WriteExcel::Chart.
#
# Copyright 2000-2009, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

require Exporter;

use strict;
use Spreadsheet::WriteExcel::Chart;


use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::WriteExcel::Chart Exporter);

$VERSION = '2.32';

###############################################################################
#
# new()
#
#
sub new {

    my $class = shift;
    my $self  = Spreadsheet::WriteExcel::Chart->new( @_ );

    bless $self, $class;
    return $self;
}


###############################################################################
#
# _store_chart_type()
#
# Implementation of the abstract method from the specific chart class.
#
# Write the LINE chart BIFF record. Defines a line chart type.
#
sub _store_chart_type {

    my $self = shift;

    my $record = 0x1018;    # Record identifier.
    my $length = 0x0002;    # Number of bytes to follow.
    my $grbit  = 0x0000;    # Option flags.

    my $header = pack 'vv', $record, $length;
    my $data = pack 'v', $grbit;

    $self->_append( $header, $data );
}


1;


__END__


=head1 NAME

Line - A writer class for Excel Line charts.

=head1 SYNOPSIS

To create a simple Excel file with a Line chart using Spreadsheet::WriteExcel:

    #!/usr/bin/perl -w

    use strict;
    use Spreadsheet::WriteExcel;

    my $workbook  = Spreadsheet::WriteExcel->new( 'chart.xls' );
    my $worksheet = $workbook->add_worksheet();

    my $chart     = $workbook->add_chart( name => 'Chart1', type => 'line' );

    # Configure the chart.
    $chart->add_series(
        categories => '=Sheet1!$A$2:$A$7',
        values     => '=Sheet1!$B$2:$B$7',
    );

    # Add the data to the worksheet the chart refers to.
    my $data = [
        [ 'Category', 2, 3, 4, 5, 6, 7 ],
        [ 'Value',    1, 4, 5, 2, 1, 5 ],
    ];

    $worksheet->write( 'A1', $data );

    __END__

=head1 DESCRIPTION

This module implements Line charts for L<Spreadsheet::WriteExcel>. The chart object is created via the Workbook C<add_chart()> method:

    my $chart = $workbook->add_chart( name => 'Chart1', type => 'line' );

Once the object is created it can be configured via the following methods that are common to all chart classes:

    $chart->add_series();
    $chart->set_x_axis();
    $chart->set_y_axis();
    $chart->set_title();

These methods are explained in detail in L<Spreadsheet::WriteExcel::Chart>. Class specific methods or settings, if any, are explained below.

=head1 Line Chart Methods

There aren't currently any line chart specific methods. See the TODO section of L<Spreadsheet::WriteExcel::Chart>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

Copyright MM-MMIX, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

