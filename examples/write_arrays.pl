#!/usr/bin/perl -w

#######################################################################
#
# Example of how to use the Spreadsheet::WriteExcel module to
# write 1D and 2D arrays of data.
#
# To find out more about array references refer(!!) to the perlref and
# perlreftut manpages. To find out more about 2D arrays or "list of
# lists" refer to the perllol manpage.
#
# reverse('©'), March 2002, John McNamara, jmcnamara@cpan.org
#


use strict;
use Spreadsheet::WriteExcel;


my $workbook   = Spreadsheet::WriteExcel->new("write_arrays.xls");
my $worksheet1 = $workbook->add_worksheet('Example 1');
my $worksheet2 = $workbook->add_worksheet('Example 2');
my $worksheet3 = $workbook->add_worksheet('Example 3');
my $worksheet4 = $workbook->add_worksheet('Example 4');
my $worksheet5 = $workbook->add_worksheet('Example 5');
my $worksheet6 = $workbook->add_worksheet('Example 6');
my $worksheet7 = $workbook->add_worksheet('Example 7');
my $worksheet8 = $workbook->add_worksheet('Example 8');

my $format     = $workbook->add_format(color => 'red', bold => 1);


# Data arrays used in the following examples.
# undef values are written as blank cells (with format if specified).
#
my @array   =   ( 'one', 'two', undef, 'four' );

my @array2d =   (
                    ['maggie', 'milly', 'molly', 'may'  ],
                    [13,       14,      15,      16     ],
                    ['shell',  'star',  'crab',  'stone'],
                );


# 1. Write a row of data using an array reference.
$worksheet1->write('A1', \@array);

# 2. Same as 1. above using an anonymous array ref.
$worksheet2->write('A1', [ @array ]);

# 3. Write a row of data using an explicit write_row() method call.
#    This is the same as calling write() in Ex. 1 above.
#
$worksheet3->write_row('A1', \@array);

# 4. Write a column of data using the write_col() method call.
$worksheet4->write_col('A1', \@array);

# 5. Write a column of data using a ref to an array ref, i.e. a 2D array.
$worksheet5->write('A1', [ \@array ]);

# 6. Write a 2D array in col-row order.
$worksheet6->write('A1', \@array2d);

# 7. Write a 2D array in row-col order.
$worksheet7->write_col('A1', \@array2d);

# 8. Write a row of data with formatting. The blank cell is also formatted.
$worksheet8->write('A1', \@array, $format);

