#!/usr/bin/perl -w

######################################################################
#
# Examples of formatting using the Spreadsheet::WriteExcel module
#
# Dec 2000, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcel;

my $workbook = Spreadsheet::WriteExcel->new("formats.xls");

# Some common formats
my $center  = $workbook->addformat();
my $heading = $workbook->addformat();
$center->set_align('center');
$heading->set_align('center');
$heading->set_bold();


# Call these subroutines to demonstrate different formatting options
intro();
fonts();
named_colors();
standard_colors();
numeric_formats();
borders();
patterns();
alignment();
misc();

# Note: this is required
$workbook->close();


######################################################################
#
# Intro.
#
sub intro {

    my $worksheet = $workbook->addworksheet('Introduction');

    $worksheet->set_column(0, 0, 60);

    my $format = $workbook->addformat();
    $format->set_bold();
    $format->set_size(14);
    $format->set_color('blue');
    $format->set_align('center');
    
    $worksheet->write(2, 0, 'This workbook demonstrates some of',  $format);
    $worksheet->write(3, 0, 'the formatting options provided by',  $format);
    $worksheet->write(4, 0, 'the Spreadsheet::WriteExcel module.', $format);
}


######################################################################
#
# Demonstrate the named colors.
#
sub named_colors {

    my @colors  = qw(aqua black blue fuchsia gray green lime navy
                     orange purple red silver white yellow);

    my @indices =   ( 0x0F, 0x08, 0x0C, 0x0E, 0x17, 0x11, 0x0B,
                      0x12, 0x1D, 0x24, 0x0A, 0x16, 0x09, 0x0D);

    my $worksheet = $workbook->addworksheet('Named colors');

    $worksheet->set_column(0, 3, 15);

    $worksheet->write(0, 0, "Index", $heading);
    $worksheet->write(0, 1, "Index", $heading);
    $worksheet->write(0, 2, "Name",  $heading);
    $worksheet->write(0, 3, "Color", $heading);

    for my $i (0..13) {
        my $format = $workbook->addformat();
        $format->set_color($colors[$i]);
        $format->set_align('center');
        
        $worksheet->write($i+1, 0, $indices[$i],                    $center);
        $worksheet->write($i+1, 1, sprintf("0x%02X", $indices[$i]), $center);
        $worksheet->write($i+1, 2, $colors[$i],                     $center);
        $worksheet->write($i+1, 3, $colors[$i],                     $format);
    }
}


######################################################################
#
# Demonstrate the standard Excel colors in the range 8..63.
#
sub standard_colors {

    my $worksheet = $workbook->addworksheet('Standard colors');

    $worksheet->set_column(0, 3, 15);

    $worksheet->write(0, 0, "Index", $heading);
    $worksheet->write(0, 1, "Index", $heading);
    $worksheet->write(0, 2, "Color", $heading);

    for my $i (8..63){
        my $format = $workbook->addformat();
        $format->set_color($i);
        $format->set_align('center');
        
        $worksheet->write(($i -7), 0, $i,                    $center);
        $worksheet->write(($i -7), 1, sprintf("0x%02X", $i), $center);
        $worksheet->write(($i -7), 2, 'COLOR',               $format);
    }
}


######################################################################
#
# Demonstrate the standard numeric formats.
#
sub numeric_formats {

    my $worksheet = $workbook->addworksheet('Numeric formats');

    $worksheet->set_column(0, 4, 15);
    $worksheet->set_column(5, 5, 45);
    
    $worksheet->write(0, 0, "Index",       $heading);
    $worksheet->write(0, 1, "Index",       $heading);
    $worksheet->write(0, 2, "Unformatted", $heading);
    $worksheet->write(0, 3, "Formatted",   $heading);
    $worksheet->write(0, 4, "Negative",    $heading);
    $worksheet->write(0, 5, "Format",      $heading);

    my @formats;
    push @formats, [ 0x00, 1234.567,   0,         'General' ];
    push @formats, [ 0x01, 1234.567,   0,         '0' ];
    push @formats, [ 0x02, 1234.567,   0,         '0.00' ];
    push @formats, [ 0x03, 1234.567,   0,         '#,##0' ];
    push @formats, [ 0x04, 1234.567,   0,         '#,##0.00' ];
    push @formats, [ 0x05, 1234.567,   -1234.567, '($#,##0_);($#,##0)' ];
    push @formats, [ 0x06, 1234.567,   -1234.567, '($#,##0_);[Red]($#,##0)' ];
    push @formats, [ 0x07, 1234.567,   -1234.567, '($#,##0.00_);($#,##0.00)' ];
    push @formats, [ 0x08, 1234.567,   -1234.567, '($#,##0.00_);[Red]($#,##0.00)' ];
    push @formats, [ 0x09, 0.567,      0,         '0%' ];
    push @formats, [ 0x0a, 0.567,      0,         '0.00%' ];
    push @formats, [ 0x0b, 1234.567,   0,         '0.00E+00' ];
    push @formats, [ 0x0c, 0.75,       0,         '# ?/?' ];
    push @formats, [ 0x0d, 0.3125,     0,         '# ??/??' ];
    push @formats, [ 0x0e, 36870.016,  0,         'm/d/yy' ];
    push @formats, [ 0x0f, 36870.016,  0,         'd-mmm-yy' ];
    push @formats, [ 0x10, 36870.016,  0,         'd-mmm' ];
    push @formats, [ 0x11, 36870.016,  0,         'mmm-yy' ];
    push @formats, [ 0x12, 36870.016,  0,         'h:mm AM/PM' ];
    push @formats, [ 0x13, 36870.016,  0,         'h:mm:ss AM/PM' ];
    push @formats, [ 0x14, 36870.016,  0,         'h:mm' ];
    push @formats, [ 0x15, 36870.016,  0,         'h:mm:ss' ];
    push @formats, [ 0x16, 36870.016,  0,         'm/d/yy h:mm' ];
    push @formats, [ 0x25, 1234.567,   -1234.567, '(#,##0_);(#,##0)' ];
    push @formats, [ 0x26, 1234.567,   -1234.567, '(#,##0_);[Red](#,##0)' ];
    push @formats, [ 0x27, 1234.567,   -1234.567, '(#,##0.00_);(#,##0.00)' ];
    push @formats, [ 0x28, 1234.567,   -1234.567, '(#,##0.00_);[Red](#,##0.00)' ];
    push @formats, [ 0x29, 1234.567,   -1234.567, '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)' ];
    push @formats, [ 0x2a, 1234.567,   -1234.567, '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)' ];
    push @formats, [ 0x2b, 1234.567,   -1234.567, '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)' ];
    push @formats, [ 0x2c, 1234.567,   -1234.567, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)' ];
    push @formats, [ 0x2d, 36870.016,  0,         'mm:ss' ];
    push @formats, [ 0x2e, 3.0153,     0,         '[h]:mm:ss' ];
    push @formats, [ 0x2f, 36870.016,  0,         'mm:ss.0' ];
    push @formats, [ 0x30, 1234.567,   0,         '##0.0E+0' ];
    push @formats, [ 0x31, 1234.567,   0,         '@' ];

    my $i; # RFC 120
    foreach my $format (@formats){
        my $style = $workbook->addformat();
        $style->set_num_format($$format[0]);

        $i++;
        $worksheet->write($i, 0, $$format[0],                    $center);
        $worksheet->write($i, 1, sprintf("0x%02X", $$format[0]), $center);
        $worksheet->write($i, 2, $$format[1],                    $center);
        $worksheet->write($i, 3, $$format[1],                    $style);

        if ($$format[2]) {
            $worksheet->write($i, 4, $$format[2], $style);
        }

        $worksheet->write_string($i, 5, $$format[3]);
    }
}


######################################################################
#
# Demonstrate the font options.
#
sub fonts {

    my $worksheet = $workbook->addworksheet('Fonts');

    $worksheet->set_column(0, 0, 30);
    $worksheet->set_column(1, 1, 10);

    $worksheet->write(0, 0, "Font name",   $heading);
    $worksheet->write(0, 1, "Font size",   $heading);

    my @fonts;
    push @fonts, [ 10, 'Arial' ];
    push @fonts, [ 12, 'Arial' ];
    push @fonts, [ 14, 'Arial' ];
    push @fonts, [ 12, 'Arial Black' ];
    push @fonts, [ 12, 'Arial Narrow' ];
    push @fonts, [ 12, 'Century Schoolbook' ];
    push @fonts, [ 12, 'Courier' ];
    push @fonts, [ 12, 'Courier New' ];
    push @fonts, [ 12, 'Garamond' ];
    push @fonts, [ 12, 'Impact' ];
    push @fonts, [ 12, 'Lucida Handwriting'] ;
    push @fonts, [ 12, 'Times New Roman' ];
    push @fonts, [ 12, 'Symbol' ];
    push @fonts, [ 12, 'Wingdings' ];
    push @fonts, [ 12, 'A font that doesnt exist' ];

    my $i;
    foreach my $font (@fonts){
        my $format = $workbook->addformat();

        $format->set_size($$font[0]);
        $format->set_font($$font[1]);

        $i++;
        $worksheet->write($i, 0, $$font[1], $format);
        $worksheet->write($i, 1, $$font[0], $format);
    }

}


######################################################################
#
# Demonstrate the standard Excel border styles.
#
sub borders {

    my $worksheet = $workbook->addworksheet('Borders');

    $worksheet->set_column(0, 4, 10);
    $worksheet->set_column(5, 5, 40);

    $worksheet->write(0, 0, "Index", $heading);
    $worksheet->write(0, 1, "Index", $heading);
    $worksheet->write(0, 3, "Style", $heading);
    $worksheet->write(0, 5, "The style is highlighted in red for ", $heading);
    $worksheet->write(1, 5, "emphasis, the default color is black.", $heading);

    for my $i (0..7){
        my $format = $workbook->addformat();
        $format->set_border($i);
        $format->set_border_color('red');
        $format->set_align('center');

        $worksheet->write((2*($i+1)), 0, $i,                    $center);
        $worksheet->write((2*($i+1)), 1, sprintf("0x%02X", $i), $center);

        $worksheet->write((2*($i+1)), 3, "Border", $format);
    }
}



######################################################################
#
# Demonstrate the standard Excel cell patterns.
#
sub patterns {

    my $worksheet = $workbook->addworksheet('Patterns');

    $worksheet->set_column(0, 4, 10);
    $worksheet->set_column(5, 5, 50);

    $worksheet->write(0, 0, "Index", $heading);
    $worksheet->write(0, 1, "Index", $heading);
    $worksheet->write(0, 3, "Pattern", $heading);

    $worksheet->write(0, 5, "The background colour has been set to silver.", $heading);
    $worksheet->write(1, 5, "The foreground colour has been set to green.",  $heading);

    for my $i (0..31){
        my $format = $workbook->addformat();

        $format->set_pattern($i);
        $format->set_bg_color('silver');
        $format->set_fg_color('green');
        $format->set_align('center');

        $worksheet->write((2*($i+1)), 0, $i,                    $center);
        $worksheet->write((2*($i+1)), 1, sprintf("0x%02X", $i), $center);

        $worksheet->write((2*($i+1)), 3, "Pattern", $format);

        if ($i == 1) {
            $worksheet->write((2*($i+1)), 5, "This is solid colour, the most useful pattern.", $heading);
        }
    }
}


######################################################################
#
# Demonstrate the standard Excel cell alignments.
#
sub alignment {

    my $worksheet = $workbook->addworksheet('Alignment');

    $worksheet->set_column(0, 7, 12);
    $worksheet->set_row(0, 40);
    $worksheet->set_selection(7, 0);

    my $format01 = $workbook->addformat();
    my $format02 = $workbook->addformat();
    my $format03 = $workbook->addformat();
    my $format04 = $workbook->addformat();
    my $format05 = $workbook->addformat();
    my $format06 = $workbook->addformat();
    my $format07 = $workbook->addformat();
    my $format08 = $workbook->addformat();
    my $format09 = $workbook->addformat();
    my $format10 = $workbook->addformat();
    my $format11 = $workbook->addformat();
    my $format12 = $workbook->addformat();
    my $format13 = $workbook->addformat();
    my $format14 = $workbook->addformat();
    my $format15 = $workbook->addformat();

    $format02->set_align('top');
    $format03->set_align('bottom');
    $format04->set_align('vcenter');
    $format05->set_align('vjustify');
    $format06->set_text_wrap();

    $format07->set_align('left');
    $format08->set_align('right');
    $format09->set_align('center');
    $format10->set_align('fill');
    $format11->set_align('justify');
    $format12->set_align('merge');

    $format13->set_rotation(1);
    $format14->set_rotation(2);
    $format15->set_rotation(3);

    $worksheet->write(0, 0, 'Vertical',   $heading);
    $worksheet->write(0, 1, 'top',        $format02);
    $worksheet->write(0, 2, 'bottom',     $format03);
    $worksheet->write(0, 3, 'vcenter',    $format04);
    $worksheet->write(0, 4, 'vjustify',   $format05);
    $worksheet->write(0, 5, "text\nwrap", $format06);

    $worksheet->write(2, 0, 'Horizontal', $heading);
    $worksheet->write(2, 1, 'left',       $format07);
    $worksheet->write(2, 2, 'right',      $format08);
    $worksheet->write(2, 3, 'center',     $format09);
    $worksheet->write(2, 4, 'fill',       $format10);
    $worksheet->write(2, 5, 'justify',    $format11);

    $worksheet->write(2, 6, 'merge',      $format12);
    $worksheet->write_blank(2, 7,         $format12);

    $worksheet->write(4, 0, 'Rotation', $heading);
    $worksheet->write(4, 1, 'Rotate 1', $format13);
    $worksheet->write(4, 2, 'Rotate 2', $format14);
    $worksheet->write(4, 3, 'Rotate 3', $format15);
}


######################################################################
#
# Demonstrate other miscellaneous features.
#
sub misc {

    my $worksheet = $workbook->addworksheet('Miscellaneous');

    $worksheet->set_column(2, 2, 25);

    my $format01 = $workbook->addformat();
    my $format02 = $workbook->addformat();
    my $format03 = $workbook->addformat();
    my $format04 = $workbook->addformat();
    my $format05 = $workbook->addformat();
    my $format06 = $workbook->addformat();
    my $format07 = $workbook->addformat();

    $format01->set_underline(0x01);
    $format02->set_underline(0x02);
    $format03->set_underline(0x21);
    $format04->set_underline(0x22);
    $format05->set_font_strikeout();
    $format06->set_font_outline();
    $format07->set_font_shadow();

    $worksheet->write(1,  2, 'Underline  0x01',          $format01);
    $worksheet->write(3,  2, 'Underline  0x02',          $format02);
    $worksheet->write(5,  2, 'Underline  0x21',          $format03);
    $worksheet->write(7,  2, 'Underline  0x22',          $format04);
    $worksheet->write(9,  2, 'Strikeout',                $format05);
    $worksheet->write(11, 2, 'Outline (Macintosh only)', $format06);
    $worksheet->write(13, 2, 'Shadow (Macintosh only)',  $format07);
}

