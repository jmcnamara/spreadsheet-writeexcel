#!/usr/bin/perl -w

###############################################################################
#
# Example of how to use the WriteExcel module to write hyperlinks
#
# Feb 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new workbook and add a worksheet
my $workbook  = Spreadsheet::WriteExcel->new("hyperlink.xls");
my $worksheet = $workbook->addworksheet();

# Format the first column
$worksheet->set_column(0, 0, 25);
$worksheet->set_selection(0, 1);


# Add a sample format
my $format = $workbook->addformat();
$format->set_size(12);
$format->set_bold();
$format->set_color('red');
$format->set_underline();


# Write some hyperlinks
$worksheet->write(0, 0, 'http://www.perl.com/'                );
$worksheet->write(1, 0, 'http://www.perl.com/', 'Perl home'   );
$worksheet->write(2, 0, 'http://www.perl.com/', undef, $format);
$worksheet->write(3, 0, 'ftp://www.perl.com/'                 );

# Write a URL that isn't a hyperlink
$worksheet->write_string(5, 0, 'http://www.perl.com/');

