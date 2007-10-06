#!/usr/bin/perl -w

##############################################################################
#
# A simple example using Spreadsheet::WriteExcel to display all available
# Unicode characters in a font.
#
# reverse('©'), May 2004, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;


my $workbook  = Spreadsheet::WriteExcel->new('unicode_list.xls');
my $worksheet = $workbook->add_worksheet();


# Set a Unicode font.
my $uni_font  = $workbook->add_format(font => 'Arial Unicode MS');

# Ascii font for labels.
my $courier   = $workbook->add_format(font => 'Courier New');


my $char = 0;

# Loop through all 32768 UTF-16BE characters.
#
for my $row (0 .. 2 ** 12 -1) {
    for my $col (0 .. 31) {

        last if $char == 0xffff;

        if ($col % 2 == 0){
            $worksheet->write_string($row, $col,
                                           sprintf('0x%04X', $char), $courier);
        }
        else {
            $worksheet->write_utf16be_string($row, $col,
                                            pack('n', $char++), $uni_font);
        }
    }
}



__END__

