#!/usr/bin/perl -w

###############################################################################
#
# This uses the Win32::OLE module to expose the Flight Simulator easter egg
# in Excel 97 SR2.
#
# reverse('©'), March 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Win32::OLE;

my $application = Win32::OLE->new("Excel.Application");
my $workbook    = $application->Workbooks->Add;
my $worksheet   = $workbook->Worksheets(1);

$application->{Visible} = 1;

$worksheet->Range("L97:X97")->Select;
$worksheet->Range("M97")->Activate;

my $message =  "Hold down Shift and Ctrl and click the ".
               "Chart Wizard icon on the toolbar.\n\n".
               "Use the mouse motion and buttons to control ".
               "movement. Try to find the monolith. ".
               "Close this dialog first.";

$application->InputBox($message);
