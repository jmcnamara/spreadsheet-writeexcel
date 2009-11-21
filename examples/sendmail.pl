#!/usr/bin/perl -w

###############################################################################
#
# Example of how to use Mail::Sender to send a Spreadsheet::WriteExcel Excel
# file as an attachment.
#
# The main thing is to ensure that you close() the Worbook before you send it.
#
# See the L<Mail::Sender> module for further details.
#
# reverse('©'), August 2002, John McNamara, jmcnamara@cpan.org
#


use strict;
use Spreadsheet::WriteExcel;
use Mail::Sender;

# Create an Excel file
my $workbook  = Spreadsheet::WriteExcel->new("sendmail.xls");
my $worksheet = $workbook->add_worksheet;

$worksheet->write('A1', "Hello World!");

$workbook->close(); # Must close before sending



# Send the file.  Change all variables to suit
my $sender = new Mail::Sender
{
    smtp => '123.123.123.123',
    from => 'Someone'
};

$sender->MailFile(
{
    to      => 'another@mail.com',
    subject => 'Excel file',
    msg     => "Here is the data.\n",
    file    => 'mail.xls',
});


