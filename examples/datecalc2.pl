#!/usr/bin/perl -w

###############################################################################
#
# Example of how to using the Date::Calc module to calculate Excel dates.
#
# NOTE: An easier way of writing dates and times is to use the newer
#       write_date_time() Worksheet method. See the date_time.pl example.
#
# reverse('©'), June 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;
use Date::Calc qw(Delta_DHMS); # You may need to install this module.


# Create a new workbook and add a worksheet
my $workbook = Spreadsheet::WriteExcel->new("excel_date2.xls");
my $worksheet = $workbook->add_worksheet();

# Expand the first column so that the date is visible.
$worksheet->set_column("A:A", 25);


# Add a format for the date
my $format =  $workbook->add_format();
$format->set_num_format('d mmmm yyy HH:MM:SS');


my $date;

# Write some dates and times
$date =  excel_date(1900, 1, 1);
$worksheet->write("A1", $date, $format);

$date =  excel_date(2000, 1, 1);
$worksheet->write("A2", $date, $format);

$date =  excel_date(2000, 4, 17, 14, 33, 15);
$worksheet->write("A3", $date, $format);


###############################################################################
#
# excel_date($years, $months, $days, $hours, $minutes, $seconds)
#
# Create an Excel date in the 1900 format. All of the arguments are optional
# but you should at least add $years.
#
# Corrects for Excel's missing leap day in 1900. See excel_time1.pl for an
# explanation.
#
sub excel_date {

    my $years   = $_[0] || 1900;
    my $months  = $_[1] || 1;
    my $days    = $_[2] || 1;
    my $hours   = $_[3] || 0;
    my $minutes = $_[4] || 0;
    my $seconds = $_[5] || 0;

    my @date = ($years, $months, $days, $hours, $minutes, $seconds);
    my @epoch = (1899, 12, 31, 0, 0, 0);

    ($days, $hours, $minutes, $seconds) = Delta_DHMS(@epoch, @date);

    my $date = $days + ($hours*3600 +$minutes*60 +$seconds)/(24*60*60);

    # Add a day for Excel's missing leap day in 1900
    $date++ if ($date > 59);

    return $date;
}

###############################################################################
#
# excel_date($years, $months, $days, $hours, $minutes, $seconds)
#
# Create an Excel date in the 1904 format. All of the arguments are optional
# but you should at least add $years.
#
# You will also need to call $workbook->set_1904() for this format to be valid.
#
sub excel_date_1904 {

    my $years   = $_[0] || 1900;
    my $months  = $_[1] || 1;
    my $days    = $_[2] || 1;
    my $hours   = $_[3] || 0;
    my $minutes = $_[4] || 0;
    my $seconds = $_[5] || 0;

    my @date = ($years, $months, $days, $hours, $minutes, $seconds);
    my @epoch = (1904, 1, 1, 0, 0, 0);

    ($days, $hours, $minutes, $seconds) = Delta_DHMS(@epoch, @date);

    my $date = $days + ($hours*3600 +$minutes*60 +$seconds)/(24*60*60);

    return $date;
}


