#!/usr/bin/perl -w

######################################################################
#
# This is a example of how to work around Spreadsheet::WriteExcel and
# Excel5's 255 character string limitation using a formula to create
# a long string from shorter strings.
#
# reverse('©'), April 2002, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;


my $workbook  = Spreadsheet::WriteExcel->new("long_string.xls");
my $worksheet = $workbook->addworksheet();


#
# The following formatting is optional.
#
my $wrap = $workbook->addformat(text_wrap => 1, valign => 'top');
$worksheet->set_column('B:B', 50);
$worksheet->set_row(1, 170);
$worksheet->set_row(2, 170);



# Example 1
#
# Create a long string using the Excel concatenation operator "&".
# The formula has the format '="String1" & "String1" & ...'
#

my $str1 = 'these are the days and ' x 10;
my $str2 = '="'. $str1 . '"&"'. $str1 . '"&"'. $str1 . 'stop."';

$worksheet->write('B2', $str2, $wrap);



# Example 2
#
# Create a long string using the Excel concatenation operator "&".
# The methodology is the same as the previous example except that
# we use a function to insert the formatting.
#

my $str3 = ('these are the days and ' x 30) . 'stop."';

$worksheet->write('B3', long_string($str3), $wrap);

# Leaves shorter strings unmodified
$worksheet->write('B4', long_string("hello, world"));




######################################################################
#
# long_string($str)
#
# Converts long strings into an Excel string concatenation formula.
# The concatenation is inserted between words to improve legibility.
#
# returns: An Excel formula if string is longer than 255 chars.
#          The unmodified string otherwise.
#
sub long_string {

    my $str   = shift;
    my $limit = 255;

    # Return short strings
    return $str if length $str <= $limit;

    # Split the line at word boundaries where possible
    my @segments = $str =~ m[.{1,$limit}$|.{1,$limit}\b|.{1,$limit}]sog;

    # Join the string back together with quotes and Excel concatenation
    $str = join '"&"', @segments;

    # Add formatting to convert the string to a formula string
    return $str = qq(="$str");
}

