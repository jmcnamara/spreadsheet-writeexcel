#!/usr/bin/perl -w

##############################################################################
#
# A simple example of writing some Unicode text with Spreadsheet::WriteExcel.
#
# This creates an Excel file with the word Nippon in 3 character sets.
#
# reverse('©'), May 2004, John McNamara, jmcnamara@cpan.org
#


use strict;
use Spreadsheet::WriteExcel;


my $workbook  = Spreadsheet::WriteExcel->new('japan.xls');
my $worksheet = $workbook->add_worksheet();


# Set a Unicode font.
my $uni_font  = $workbook->add_format(font => 'Arial Unicode MS');


# Create some UTF-16BE Unicode text.
my $kanji     = pack 'n*', 0x65e5, 0x672c;
my $katakana  = pack 'n*', 0xff86, 0xff8e, 0xff9d;
my $hiragana  = pack 'n*', 0x306b, 0x307b, 0x3093;



$worksheet->write_unicode('A1', $kanji,    $uni_font);
$worksheet->write_unicode('A2', $katakana, $uni_font);
$worksheet->write_unicode('A3', $hiragana, $uni_font);


$worksheet->write('B1', 'Kanji');
$worksheet->write('B2', 'Katakana');
$worksheet->write('B3', 'Hiragana');


__END__


