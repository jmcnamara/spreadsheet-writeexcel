package Spreadsheet::WriteExcel::Chart::Column;

###############################################################################
#
# Column - A writer class for Excel Column charts.
#
#
# Used in conjunction with Spreadsheet::WriteExcel
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

$VERSION = '2.31';

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


1;


__END__


=head1 NAME

Column - A writer class for Excel Column charts.

=head1 SYNOPSIS

See the documentation for Spreadsheet::WriteExcel::Chart

=head1 DESCRIPTION

This module is used in conjunction with Spreadsheet::WriteExcel.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

Copyright MM-MMIX, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

