#!/usr/bin/perl -w

##############################################################################
#
# A simple example of writing some Unicode text with Spreadsheet::WriteExcel.
#
# reverse('©'), May 2004, John McNamara, jmcnamara@cpan.org
#


use strict;
use Spreadsheet::WriteExcel;


my $workbook  = Spreadsheet::WriteExcel->new('unicode.xls');
my $worksheet = $workbook->add_worksheet();


# Write the Unicode smiley face (with increased font for legibility)
my $smiley    = pack "n", 0x263a;
my $big_font  = $workbook->add_format(size => 40);

$worksheet->write_unicode('A3', $smiley, $big_font);


# Write a phrase in Cyrillic
my $uni_str = pack "H*", "042d0442043e002004440440043004370430002004".
                         "3d043000200440044304410441043a043e043c0021";

$worksheet->write_unicode('A5', $uni_str);


$worksheet->write_unicode('A7', pack "H*", "0074006500730074");





__END__

