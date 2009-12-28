#!/usr/bon/perl -w

#######################################################################
#
# A simple demo of a Column chart in Spreadsheet::WriteExcel.
#
# reverse('©'), October 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook  = Spreadsheet::WriteExcel->new( 'chart_column.xls' );
my $worksheet = $workbook->add_worksheet();
my $chart     = $workbook->add_chart( name => 'Chart1', type => 'column' );

# Configure the chart.
$chart->add_series(
    series        => '=Sheet1!$A$1:$A$10',
    name          => 'Batch 1',
);

$chart->add_series(
    series        => '=Sheet1!$B$1:$B$10',
    name          => 'Batch 2',
);
$chart->set_x_axis( name => 'Sample (number)', );
$chart->set_y_axis( name => 'Weight (kg)', );
$chart->set_title ( name => 'Some sample test data' );

# Add the data the the chart refers to.
my $data = [
    [2, 3, 6, 7, 5, 4, 8, 1, 5, 4],
    [6, 7, 5, 2, 1, 1, 3, 5, 4, 1],
];

$worksheet->write( 'A1', $data );

__END__

