package Spreadsheet::WriteExcel::Big;

###############################################################################
#
# Used in conjuction with the big.pl program. Change the extension to .pm
# and copy the file to the Spreadsheet/WriteExcel directory.
#


###############################################################################
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
# Documentation after __END__
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
