package Spreadsheet::WriteExcel::Big;

###############################################################################
#
# Used in conjunction with the big.pl program.
#
#
# Example of how to use extend the  Spreadsheet::WriteExcel 7MB limit with
# OLE::Storage_Lite http://search.cpan.org/search?dist=OLE-Storage_Lite
#
# Nov 2000, Kawai, Takanori (Hippo2000)
#   Mail: GCD00051@nifty.ne.jp
#   http://member.nifty.ne.jp/hippo2000


###############################################################################
#
# WriteExcel::Big
#
# Spreadsheet::WriteExcel - Write formatted text and numbers to a
# cross-platform Excel binary file.
#
# Copyright 2000, John McNamara, jmcnamara@cpan.org
#
#

require Exporter;

use strict;
use Spreadsheet::WriteExcel::WorkbookBig;




use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::WriteExcel::WorkbookBig Exporter);

$VERSION = '0.23'; # 10 December 2000

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


=head1 AUTHOR


John McNamara jmcnamara@cpan.org


=head1 COPYRIGHT


© MM-MMI, John McNamara.


All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.