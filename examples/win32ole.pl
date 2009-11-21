#!/usr/bin/perl -w

###############################################################################
#
# This is a simple example of how to create an Excel file using the
# Win32::OLE module for the sake of comparison.
#
# reverse('©'), March 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Cwd;
use Win32::OLE;
use Win32::OLE::Const 'Microsoft Excel';


my $application = Win32::OLE->new("Excel.Application");
my $workbook    = $application->Workbooks->Add;
my $worksheet   = $workbook->Worksheets(1);

$worksheet->Cells(1,1)->{Value} = "Hello World";
$worksheet->Cells(2,1)->{Value} = "One";
$worksheet->Cells(3,1)->{Value} = "Two";
$worksheet->Cells(4,1)->{Value} =  3;
$worksheet->Cells(5,1)->{Value} =  4.0000001;

# Add some formatting
$worksheet->Cells(1,1)->Font->{Bold}       = "True";
$worksheet->Cells(1,1)->Font->{Size}       = 16;
$worksheet->Cells(1,1)->Font->{ColorIndex} = 3;
$worksheet->Columns("A:A")->{ColumnWidth}  = 25;

# Write a hyperlink
my $range = $worksheet->Range("A7:A7");
$worksheet->Hyperlinks->Add({ Anchor => $range, Address => "http://www.perl.com/"});

# Get current directory using Cwd.pm
my $dir = cwd();

$workbook->SaveAs({
                    FileName   => $dir . '/win32ole.xls',
                    FileFormat => xlNormal,
                  });
$workbook->Close;
