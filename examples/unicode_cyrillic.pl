#!/usr/bin/perl -w

##############################################################################
#
# A simple example of writing some Russian cyrillic text using
# Spreadsheet::WriteExcel and perl 5.8.
#
# reverse('©'), March 2005, John McNamara, jmcnamara@cpan.org
#



# Perl 5.8 or later is required for proper utf8 handling. For older perl
# versions you should use UTF16 and the write_utf16be_string() method.
# See the write_utf16be_string section of the Spreadsheet::WriteExcel docs.
#
require 5.008;

use strict;
use Spreadsheet::WriteExcel;


# In this example we generate utf8 strings from character data but in a
# real application we would expect them to come from an external source.
#


# Create a Russian worksheet name in utf8.
my $sheet   = pack "U*", 0x0421, 0x0442, 0x0440, 0x0430, 0x043D, 0x0438,
                         0x0446, 0x0430;


# Create a Russian string.
my $str     = pack "U*", 0x0417, 0x0434, 0x0440, 0x0430, 0x0432, 0x0441,
                         0x0442, 0x0432, 0x0443, 0x0439, 0x0020, 0x041C,
                         0x0438, 0x0440, 0x0021;



my $workbook  = Spreadsheet::WriteExcel->new("unicode_cyrillic.xls");
my $worksheet = $workbook->add_worksheet($sheet . '1');

   $worksheet->set_column('A:A', 18);
   $worksheet->write('A1', $str);


__END__

