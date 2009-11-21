#!/usr/bin/perl -w


######################################################################
#
# NOTE: An easier way of writing dates and times is to use the newer
#       write_date_time() Worksheet method. See the date_time.pl example.
#
######################################################################
#
# Demonstration of writing date/time cells to Excel spreadsheets,
# using UNIX/Perl time as source of date/time.
#
######################################################################
#
# UNIX/Perl time is the time since the Epoch (00:00:00 GMT, 1 Jan 1970)
# measured in seconds.
#
# An Excel file can use exactly one of two different date/time systems.
# In these systems, a floating point number represents the number of days
# (and fractional parts of the day) since a start point. The floating point
# number is referred to as a 'serial'.
#
# The two systems ('1900' and '1904') use different starting points:
#
#  '1900'; '1.00' is 1 Jan 1900 BUT 1900 is erroneously regarded as
#          a leap year - see:
#            http://support.microsoft.com/support/kb/articles/Q181/3/70.asp
#          for the excuse^H^H^H^H^H^Hreason.
#  '1904'; '1.00' is 2 Jan 1904.
#
# The '1904' system is the default for Apple Macs. Windows versions of
# Excel have the option to use the '1904' system.
#
# Note that Visual Basic's "DateSerial" function does NOT erroneously
# regard 1900 as a leap year, and thus its serials do not agree with
# the 1900 serials of Excel for dates before 1 Mar 1900.
#
# Note that StarOffice (at least at version 5.2) does NOT erroneously
# regard 1900 as a leap year, and thus its serials do not agree with
# the 1900 serials of Excel for dates before 1 Mar 1900.
#

# Copyright 2000, Andrew Benham, adsb@bigfoot.com
#

######################################################################
#
# Calculation description
# =======================
#
# 1900 system
# -----------
# Unix time is '0' at 00:00:00 GMT 1 Jan 1970, i.e. 70 years after 1 Jan 1900.
# Of those 70 years, 17 (1904,08,12,16,20,24,28,32,36,40,44,48,52,56,60,64,68)
# were leap years with an extra day.
# Thus there were 17 + 70*365 days = 25567 days between 1 Jan 1900 and
# 1 Jan 1970.
# In the 1900 system, '1' is 1 Jan 1900, but as 1900 was not a leap year
# 1 Jan 1900 should really be '2', so 1 Jan 1970 is '25569'.
#
# 1904 system
# -----------
# Unix time is '0' at 00:00:00 GMT 1 Jan 1970, i.e. 66 years after 1 Jan 1904.
# Of those 66 years, 17 (1904,08,12,16,20,24,28,32,36,40,44,48,52,56,60,64,68)
# were leap years with an extra day.
# Thus there were 17 + 66*365 days = 24107 days between 1 Jan 1904 and
# 1 Jan 1970.
# In the 1904 system, 2 Jan 1904 being '1', 1 Jan 1970 is '24107'.
#
######################################################################
#
# Copyright (c) 2000, Andrew Benham.
# This program is free software. It may be used, redistributed and/or
# modified under the same terms as Perl itself.
#
# Andrew Benham, adsb@bigfoot.com
# London, United Kingdom
# 11 Nov 2000
#
######################################################################


use strict;
use Spreadsheet::WriteExcel;

use Time::Local;

use vars qw/$DATE_SYSTEM/;

# Use 1900 date system on all platforms other than Apple Mac (for which
# use 1904 date system).
$DATE_SYSTEM = ($^O eq 'MacOS') ? 1 : 0;

my $workbook = Spreadsheet::WriteExcel->new("dates.xls");
my $worksheet = $workbook->add_worksheet();

my $format_date =  $workbook->add_format();
$format_date->set_num_format('d mmmm yyy');

$worksheet->set_column(0,1,21);

$worksheet->write_string (0,0,"The epoch (GMT)");
$worksheet->write_number (0,1,&calc_serial(0,1),0x16);

$worksheet->write_string (1,0,"The epoch (localtime)");
$worksheet->write_number (1,1,&calc_serial(0,0),0x16);

$worksheet->write_string (2,0,"Today");
$worksheet->write_number (2,1,&calc_serial(),$format_date);

my $christmas2000 = timelocal(0,0,0,25,11,100);
$worksheet->write_string (3,0,"Christmas 2000");
$worksheet->write_number (3,1,&calc_serial($christmas2000),$format_date);

$workbook->close();

#-----------------------------------------------------------
# calc_serial()
#
# Called with (up to) 2 parameters.
#   1.  Unix timestamp.  If omitted, uses current time.
#   2.  GMT flag. Set to '1' to return serial in GMT.
#       If omitted, returns serial in appropriate timezone.
#
# Returns date/time serial according to $DATE_SYSTEM selected
#-----------------------------------------------------------
sub calc_serial {
	my $time = (defined $_[0]) ? $_[0] : time();
	my $gmtflag = (defined $_[1]) ? $_[1] : 0;

	# Divide timestamp by number of seconds in a day.
	# This gives a date serial with '0' on 1 Jan 1970.
	my $serial = $time / 86400;

	# Adjust the date serial by the offset appropriate to the
	# currently selected system (1900/1904).
	if ($DATE_SYSTEM == 0) {	# use 1900 system
		$serial += 25569;
	} else {			# use 1904 system
		$serial += 24107;
	}

	unless ($gmtflag) {
		# Now have a 'raw' serial with the right offset. But this
		# gives a serial in GMT, which is false unless the timezone
		# is GMT. We need to adjust the serial by the appropriate
		# timezone offset.
		# Calculate the appropriate timezone offset by seeing what
		# the differences between localtime and gmtime for the given
		# time are.

		my @gmtime = gmtime($time);
		my @ltime  = localtime($time);

		# For the first 7 elements of the two arrays, adjust the
		# date serial where the elements differ.
		for (0 .. 6) {
			my $diff = $ltime[$_] - $gmtime[$_];
			if ($diff) {
				$serial += _adjustment($diff,$_);
			}
		}
	}

	# Perpetuate the error that 1900 was a leap year by decrementing
	# the serial if we're using the 1900 system and the date is prior to
	# 1 Mar 1900. This has the effect of making serial value '60'
	# 29 Feb 1900.

	# This fix only has any effect if UNIX/Perl time on the platform
	# can represent 1900. Many can't.

	unless ($DATE_SYSTEM) {
		$serial-- if ($serial < 61);	# '61' is 1 Mar 1900
	}
	return $serial;
}

sub _adjustment {
	# Based on the difference in the localtime/gmtime array elements
	# number, return the adjustment required to the serial.

	# We only look at some elements of the localtime/gmtime arrays:
	#    seconds    unlikely to be different as all known timezones
	#               have an offset of integral multiples of 15 minutes,
	#		but it's easy to do.
	#    minutes    will be different for timezone offsets which are
	#		not an exact number of hours.
	#    hours	very likely to be different.
	#    weekday	will differ when localtime/gmtime difference
	#		straddles midnight.
	#
	# Assume that difference between localtime and gmtime is less than
	# 5 days, then don't have to do maths for day of month, month number,
	# year number, etc...

	my ($delta,$element) = @_;
	my $adjust = 0;

	if ($element == 0) {		# Seconds
		$adjust = $delta/86400;		# 60 * 60 * 24
	} elsif ($element == 1) {	# Minutes
		$adjust = $delta/1440;		# 60 * 24
	} elsif ($element == 2) {	# Hours
		$adjust = $delta/24;		# 24
	} elsif ($element == 6) {	# Day of week number
		# Catch difference straddling Sat/Sun in either direction
		$delta += 7 if ($delta < -4);
		$delta -= 7 if ($delta > 4);

		$adjust = $delta;
	}
	return $adjust;
}

