#!/usr/bin/perl -w

#######################################################################
#
# Example of how to insert images into an Excel worksheet using the
# Spreadsheet::WriteExcel insert_image() method.
#
# reverse('©'), October 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

# Create a new workbook called simple.xls and add a worksheet
my $workbook   = Spreadsheet::WriteExcel->new("images.xls");
my $worksheet1 = $workbook->add_worksheet('Image 1');
my $worksheet2 = $workbook->add_worksheet('Image 2');
my $worksheet3 = $workbook->add_worksheet('Image 3');
my $worksheet4 = $workbook->add_worksheet('Image 4');

# Insert a basic image
$worksheet1->write('A10', "Image inserted into worksheet.");
$worksheet1->insert_image('A1', 'republic.png');


# Insert an image with an offset
$worksheet2->write('A10', "Image inserted with an offset.");
$worksheet2->insert_image('A1', 'republic.png', 32, 10);

# Insert a scaled image
$worksheet3->write('A10', "Image scaled: width x 2, height x 0.8.");
$worksheet3->insert_image('A1', 'republic.png', 0, 0, 2, 0.8);

# Insert an image over varied column and row sizes.
$worksheet4->set_column('A:A', 5);
$worksheet4->set_column('B:B', undef, undef, 1); # Hidden
$worksheet4->set_column('C:D', 10);
$worksheet4->set_row(0, 30);
$worksheet4->set_row(3, 5);

$worksheet4->write('A10', "Image inserted over scaled rows and columns.");
$worksheet4->insert_image('A1', 'republic.png');




