#!/usr/bin/perl -w

###############################################################################
#
# Example of using Spreadsheet::WriteExcel to write Excel files to
# different filehandles.
#
# reverse('©'), April 2003, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;
use IO::Scalar;




###############################################################################
#
# Example 1. This demonstrates the standard way of creating an Excel file by
# specifying a file name.
#

my $workbook1  = Spreadsheet::WriteExcel->new('fh_01.xls');
my $worksheet1 = $workbook1->add_worksheet();

$worksheet1->write(0, 0,  "Hi Excel!");




###############################################################################
#
# Example 2. Write an Excel file to an existing filehandle.
#

open    TEST, "> fh_02.xls" or die "Couldn't open file: $!";
binmode TEST; # Always do this regardless of whether the platform requires it.

my $workbook2  = Spreadsheet::WriteExcel->new(\*TEST);
my $worksheet2 = $workbook2->add_worksheet();

$worksheet2->write(0, 0,  "Hi Excel!");




###############################################################################
#
# Example 3. Write an Excel file to an existing OO style filehandle.
#

my $fh = FileHandle->new("> fh_03.xls")
         or die "Couldn't open file: $!";

binmode($fh);

my $workbook3  = Spreadsheet::WriteExcel->new($fh);
my $worksheet3 = $workbook3->add_worksheet();

$worksheet3->write(0, 0,  "Hi Excel!");




###############################################################################
#
# Example 4. Write an Excel file to a string via IO::Scalar. Please refer to
# the IO::Scalar documentation for further details.
#

my $xls_str;

tie *XLS, 'IO::Scalar', \$xls_str;

my $workbook4  = Spreadsheet::WriteExcel->new(\*XLS);
my $worksheet4 = $workbook4->add_worksheet();

$worksheet4->write(0, 0, "Hi Excel 4");
$workbook4->close(); # This is required before we use the scalar


# The Excel file is now in $xls_str. As a demonstration, print it to a file.
open    TMP, "> fh_04.xls" or die "Couldn't open file: $!";
binmode TMP;
print   TMP  $xls_str;
close   TMP;




###############################################################################
#
# Example 5. Write an Excel file to a string via IO::Scalar's newer interface.
# Please refer to the IO::Scalar documentation for further details.
#
my $xls_str2;

my $fh5 = IO::Scalar->new(\$xls_str2);


my $workbook5  = Spreadsheet::WriteExcel->new($fh5);
my $worksheet5 = $workbook5->add_worksheet();

$worksheet5->write(0, 0, "Hi Excel 5");
$workbook5->close(); # This is required before we use the scalar

# The Excel file is now in $xls_str. As a demonstration, print it to a file.
open    TMP, "> fh_05.xls" or die "Couldn't open file: $!";
binmode TMP;
print   TMP  $xls_str2;
close   TMP;


