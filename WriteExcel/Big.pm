package Spreadsheet::WriteExcel::Big;


###############################################################################
#
# WriteExcel::Big
#
# Spreadsheet::WriteExcel - Write formatted text and numbers to a
# cross-platform Excel binary file.
#
# © MM-MMIII, John McNamara.
#
#

require Exporter;

use strict;
use Spreadsheet::WriteExcel::WorkbookBig;




use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::WriteExcel::WorkbookBig Exporter);

$VERSION = '0.32'; # May 2000

###############################################################################
#
# new()
#
# Constructor. Wrapper for a Workbook object.
# uses: Spreadsheet::WriteExcel::BIFFwriter
#       Spreadsheet::WriteExcel::OLEwriter
#       Spreadsheet::WriteExcel::WorkbookBig
#       Spreadsheet::WriteExcel::Worksheet
#       Spreadsheet::WriteExcel::Format
#
sub new {

    my $class = shift;
    my $self  = Spreadsheet::WriteExcel::WorkbookBig->new($_[0]);

    bless  $self, $class;
    return $self;
}


1;


__END__



=head1 NAME


Big - A class for creating Excel files > 7MB.


=head1 SYNOPSIS


See the documentation for Spreadsheet::WriteExcel.


=head1 DESCRIPTION


This module is used in conjunction with Spreadsheet::WriteExcel.

    use Spreadsheet::WriteExcel::Big;

    my $workbook  = Spreadsheet::WriteExcel::Big->new("file.xls");
    my $worksheet = $workbook->add_worksheet();

    # Same as Spreadsheet::WriteExcel
    ...
    ...


=head1 REQUIREMENTS

IO::Stringy and OLE::Storage_Lite


=head1 AUTHOR


John McNamara jmcnamara@cpan.org


=head1 COPYRIGHT


© MM-MMIII, John McNamara.


All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
