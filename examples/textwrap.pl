#!/usr/bin/perl -w

###############################################################################
#
# Example of formatting using the Spreadsheet::WriteExcel module
#
# This example shows how to wrap text in a cell. There are two alternatives,
# vertical justification and text wrap.
#
# With vertical justification the text is wrapped automatically to fit the
# column width. With text wrap you must specify a newline with an embedded \n.
#
# reverse('©'), March 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new workbook and add a worksheet
my $workbook  = Spreadsheet::WriteExcel->new("textwrap.xls");
my $worksheet = $workbook->add_worksheet();

# Set the column width for columns 1, 2 and 3
$worksheet->set_column(1, 1, 24);
$worksheet->set_column(2, 2, 34);
$worksheet->set_column(3, 3, 34);

# Set the row height for rows 1, 4, and 6. The height of row 2 will adjust
# automatically to fit the text.
#
$worksheet->set_row(0, 30);
$worksheet->set_row(3, 40);
$worksheet->set_row(5, 80);


# No newlines
my $str1  = "For whatever we lose (like a you or a me) ";
$str1    .= "it's always ourselves we find in the sea";

# Embedded newlines
my $str2  = "For whatever we lose\n(like a you or a me)\n";
   $str2 .= "it's always ourselves\nwe find in the sea";


# Create a format for the column headings
my $header = $workbook->add_format();
$header->set_bold();
$header->set_font("Courier New");
$header->set_align('center');
$header->set_align('vcenter');

# Create a "vertical justification" format
my $format1 = $workbook->add_format();
$format1->set_align('vjustify');

# Create a "text wrap" format
my $format2 = $workbook->add_format();
$format2->set_text_wrap();

# Write the headers
$worksheet->write(0, 1, "set_align('vjustify')", $header);
$worksheet->write(0, 2, "set_align('vjustify')", $header);
$worksheet->write(0, 3, "set_text_wrap()", $header);

# Write some examples
$worksheet->write(1, 1, $str1, $format1);
$worksheet->write(1, 2, $str1, $format1);
$worksheet->write(1, 3, $str2, $format2);

$worksheet->write(3, 1, $str1, $format1);
$worksheet->write(3, 2, $str1, $format1);
$worksheet->write(3, 3, $str2, $format2);

$worksheet->write(5, 1, $str1, $format1);
$worksheet->write(5, 2, $str1, $format1);
$worksheet->write(5, 3, $str2, $format2);




