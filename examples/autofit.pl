#!/usr/bin/perl -w

##############################################################################
#
# Simulate Excel's autofit for column widths.
#
# Excel provides a function called Autofit (Format->Columns->Autofit) that
# adjusts column widths to match the length of the longest string in a column.
# Excel calculates these widths at run time when it has access to information
# about string lengths and font information. This function is *not* a feature
# of the file format and thus cannot be implemented by Spreadsheet::WriteExcel.
#
# However, we can make an attempt to simulate it by keeping track of the
# longest string written to each column and then adjusting the column widths
# prior to closing the file.
#
# We keep track of the longest strings by adding a handler to the write()
# function. See add_handler() in the S::WE docs for more information.
#
# The main problem with trying to simulate Autofit lies in defining a
# relationship between a string length and its width in a arbitrary font and
# size. We use two approaches below. The first is a simple direct relationship
# obtained by trial and error. The second is a slightly more sophisticated
# method using an external module. For more complicated applications you will
# probably have to work out your own methods.
#
# reverse('©'), May 2006, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook    = Spreadsheet::WriteExcel->new('autofit.xls');
my $worksheet   = $workbook->add_worksheet();


###############################################################################
#
# Add a handler to store the width of the longest string written to a column.
# We use the stored width to simulate an autofit of the column widths.
#
# You should do this for every worksheet you want to autofit.
#
$worksheet->add_write_handler(qr[\w], \&store_string_widths);



$worksheet->write('A1', 'Hello');
$worksheet->write('B1', 'Hello World');
$worksheet->write('D1', 'Hello');
$worksheet->write('F1', 'This is a long string as an example.');

# Run the autofit after you have finished writing strings to the workbook.
autofit_columns($worksheet);



###############################################################################
#
# Functions used for Autofit.
#
###############################################################################

###############################################################################
#
# Adjust the column widths to fit the longest string in the column.
#
sub autofit_columns {

    my $worksheet = shift;
    my $col       = 0;

    for my $width (@{$worksheet->{__col_widths}}) {

        $worksheet->set_column($col, $col, $width) if $width;
        $col++;
    }
}


###############################################################################
#
# The following function is a callback that was added via add_write_handler()
# above. It modifies the write() function so that it stores the maximum
# unwrapped width of a string in a column.
#
sub store_string_widths {

    my $worksheet = shift;
    my $col       = $_[1];
    my $token     = $_[2];

    # Ignore some tokens that we aren't interested in.
    return if not defined $token;       # Ignore undefs.
    return if $token eq '';             # Ignore blank cells.
    return if ref $token eq 'ARRAY';    # Ignore array refs.
    return if $token =~ /^=/;           # Ignore formula

    # Ignore numbers
    return if $token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/;

    # Ignore various internal and external hyperlinks. In a real scenario
    # you may wish to track the length of the optional strings used with
    # urls.
    return if $token =~ m{^[fh]tt?ps?://};
    return if $token =~ m{^mailto:};
    return if $token =~ m{^(?:in|ex)ternal:};


    # We store the string width as data in the Worksheet object. We use
    # a double underscore key name to avoid conflicts with future names.
    #
    my $old_width    = $worksheet->{__col_widths}->[$col];
    my $string_width = string_width($token);

    if (not defined $old_width or $string_width > $old_width) {
        # You may wish to set a minimum column width as follows.
        #return undef if $string_width < 10;

        $worksheet->{__col_widths}->[$col] = $string_width;
    }


    # Return control to write();
    return undef;
}


###############################################################################
#
# Very simple conversion between string length and string width for Arial 10.
# See below for a more sophisticated method.
#
sub string_width {

    return 0.9 * length $_[0];
}

__END__



###############################################################################
#
# This function uses an external module to get a more accurate width for a
# string. Note that in a real program you could "use" the module instead of
# "require"-ing it and you could make the Font object global to avoid repeated
# initialisation.
#
# Note also that the $pixel_width to $cell_width is specific to Arial. For
# other fonts you should calculate appropriate relationships. A future version
# of S::WE will provide a way of specifying column widths in pixels instead of
# cell units in order to simplify this conversion.
#
sub string_width {

    require Font::TTFMetrics;

    my $arial        = Font::TTFMetrics->new('c:\windows\fonts\arial.ttf');

    my $font_size    = 10;
    my $dpi          = 96;
    my $units_per_em = $arial->get_units_per_em();
    my $font_width   = $arial->string_width($_[0]);

    # Convert to pixels as per TTFMetrics docs.
    my $pixel_width  = 6 + $font_width *$font_size *$dpi /(72 *$units_per_em);

    # Add extra pixels for border around text.
    $pixel_width  += 6;

    # Convert to cell width (for Arial) and for cell widths > 1.
    my $cell_width   = ($pixel_width -5) /7;

    return $cell_width;

}

__END__

