#!/usr/bin/perl -w

######################################################################
#
# Example of writing repeated formulas.
#
# reverse('©'), August 2002, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook  = Spreadsheet::WriteExcel->new("repeat.xls");
my $worksheet = $workbook->add_worksheet();


my $limit = 1000;

# Write a column of numbers
for my $row (0..$limit) {
    $worksheet->write($row, 0,  $row);
}


# Store a formula
my $formula = $worksheet->store_formula('=A1*5+4');


# Write a column of formulas based on the stored formula
for my $row (0..$limit) {
    $worksheet->repeat_formula($row, 1, $formula, undef,
                                        qr/^A1$/, 'A'.($row+1));
}


# Direct formula writing. As a speed comparison uncomment the
# following and run the program again

#for my $row (0..$limit) {
#    $worksheet->write_formula($row, 2, '=A'.($row+1).'*5+4');
#}



__END__

