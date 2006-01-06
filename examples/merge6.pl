#!/usr/bin/perl -w

###############################################################################
#
# Example of how to use the Spreadsheet::WriteExcel merge_cells() workbook
# method with Unicode strings.
#
#
# reverse('©'), December 2005, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new workbook and add a worksheet
my $workbook  = Spreadsheet::WriteExcel->new('merge6.xls');
my $worksheet = $workbook->add_worksheet();


# Increase the cell size of the merged cells to highlight the formatting.
$worksheet->set_row($_, 36) for 2..9;
$worksheet->set_column('B:D', 25);


# Format for the merged cells.
my $format = $workbook->add_format(
                                    border      => 6,
                                    bold        => 1,
                                    color       => 'red',
                                    size        => 20,
                                    valign      => 'vcentre',
                                    align       => 'left',
                                    indent      => 1,
                                  );




###############################################################################
#
# Write an Ascii string.
#

$worksheet->merge_range('B3:D4', 'ASCII: A simple string', $format);




###############################################################################
#
# Write a UTF-16 Unicode string.
#

# A phrase in Cyrillic encoded as UTF-16BE.
my $utf16_str = pack "H*", '005500540046002d00310036003a0020'.
                           '042d0442043e002004440440043004370430002004'.
                           '3d043000200440044304410441043a043e043c0021';

# Note the extra parameter at the end to indicate UTF-16 encoding.
$worksheet->merge_range('B6:D7', $utf16_str, $format, 1);




###############################################################################
#
# Write a UTF-8 Unicode string.
#

if ($] >= 5.008) {
    my $smiley = chr 0x263a;
    $worksheet->merge_range('B9:D10', "UTF-8: A Unicode smiley $smiley",
                                       $format);
}
else {
    $worksheet->merge_range('B9:D10', "UTF-8: Requires Perl 5.8", $format);
}




__END__
