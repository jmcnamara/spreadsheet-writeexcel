package Spreadsheet::WriteExcel::Format;

###############################################################################
#
# Format - A class for defining Excel formatting.
#
#
# Used in conjunction with Spreadsheet::WriteExcel
#
# Copyright 2000-2003, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

use Exporter;
use strict;







use vars qw($AUTOLOAD $VERSION @ISA);
@ISA = qw(Exporter);

$VERSION = '0.08';

###############################################################################
#
# new()
#
# Constructor
#
sub new {

    my $class  = shift;

    my $self   = {
                    _xf_index       => shift || 0,

                    _font_index     => 0,
                    _font           => 'Arial',
                    _size           => 10,
                    _bold           => 0x0190,
                    _italic         => 0,
                    _color          => 0x7FFF,
                    _underline      => 0,
                    _font_strikeout => 0,
                    _font_outline   => 0,
                    _font_shadow    => 0,
                    _font_script    => 0,
                    _font_family    => 0,
                    _font_charset   => 0,

                    _num_format     => 0,

                    _hidden         => 0,
                    _locked         => 1,

                    _text_h_align   => 0,
                    _text_wrap      => 0,
                    _text_v_align   => 2,
                    _text_justlast  => 0,
                    _rotation       => 0,

                    _fg_color       => 0x40,
                    _bg_color       => 0x41,

                    _pattern        => 0,

                    _bottom         => 0,
                    _top            => 0,
                    _left           => 0,
                    _right          => 0,

                    _bottom_color   => 0x40,
                    _top_color      => 0x40,
                    _left_color     => 0x40,
                    _right_color    => 0x40,

                    _merge_range    => 0,
                 };

    bless  $self, $class;

    # Set properties passed to Workbook::add_format()
    $self->set_properties(@_) if @_;

    return $self;
}


###############################################################################
#
# copy($format)
#
# Copy the attributes of another Spreadsheet::WriteExcel::Format object.
#
sub copy {
    my $self  = shift;
    my $other = $_[0];

    return unless defined $other;
    return unless (ref($self) eq ref($other));

    my $xf = $self->{_xf_index};    # Store XF index assigned by Workbook.pm
    %$self = %$other;               # Copy properties
    $self->{_xf_index} = $xf;       # Restore XF index
}


###############################################################################
#
# get_xf($style)
#
# Generate an Excel BIFF XF record.
#
sub get_xf {

    use integer;    # Avoid << shift bug in Perl 5.6.0 on HP-UX

    my $self      = shift;

    my $record;     # Record identifier
    my $length;     # Number of bytes to follow

    my $ifnt;       # Index to FONT record
    my $ifmt;       # Index to FORMAT record
    my $style;      # Style and other options
    my $align;      # Alignment
    my $icv;        # fg and bg pattern colors
    my $fill;       # Fill and border line style
    my $border1;    # Border line style and color
    my $border2;    # Border color


    # Set the type of the XF record and some of the attributes.
    if ($_[0] eq "style") {
        $style = 0xFFF5;
    }
    else {
        $style   = $self->{_locked};
        $style  |= $self->{_hidden} << 1;
    }


    # Flags to indicate if attributes have been set.
    my $atr_num     = ($self->{_num_format}     != 0);

    my $atr_fnt     = ($self->{_font_index}     != 0);

    my $atr_alc     = ($self->{_text_h_align}   != 0  ||
                       $self->{_text_v_align}   != 2  ||
                       $self->{_text_wrap}      != 0) ? 1 : 0;

    my $atr_bdr     = ($self->{_bottom}         != 0  ||
                       $self->{_top}            != 0  ||
                       $self->{_left}           != 0  ||
                       $self->{_right}          != 0) ? 1: 0;

    my $atr_pat     = ($self->{_fg_color}       != 0x40  ||
                       $self->{_bg_color}       != 0x41  ||
                       $self->{_pattern}        != 0x00) ? 1 : 0;

    my $atr_prot    = ($self->{_hidden}         != 0  ||
                       $self->{_locked}         != 1) ? 1 : 0;


    # Reset the default colours for the non-font properties
    $self->{_fg_color}     = 0x40 if $self->{_fg_color}     == 0x7FFF;
    $self->{_bg_color}     = 0x41 if $self->{_bg_color}     == 0x7FFF;
    $self->{_bottom_color} = 0x40 if $self->{_bottom_color} == 0x7FFF;
    $self->{_top_color}    = 0x40 if $self->{_top_color}    == 0x7FFF;
    $self->{_left_color}   = 0x40 if $self->{_left_color}   == 0x7FFF;
    $self->{_right_color}  = 0x40 if $self->{_right_color}  == 0x7FFF;


    # Zero the default border colour if the border has not been set.
    $self->{_bottom_color} = 0 if $self->{_bottom} == 0;
    $self->{_top_color}    = 0 if $self->{_top}    == 0;
    $self->{_right_color}  = 0 if $self->{_right}  == 0;
    $self->{_left_color}   = 0 if $self->{_left}   == 0;


    # The following 2 logical statements take care of special cases in relation
    # to cell colours and patterns:
    # 1. For a solid fill (_pattern == 1) Excel reverses the role of foreground
    #    and background colours.
    # 2. If the user specifies a foreground or background colour without a
    #    pattern they probably wanted a solid fill, so we fill in the defaults.
    #
    if ($self->{_pattern}  <= 0x01 and
        $self->{_bg_color} != 0x41 and
        $self->{_fg_color} == 0x40    )
    {
        $self->{_fg_color} = $self->{_bg_color};
        $self->{_bg_color} = 0x40;
        $self->{_pattern}  = 1;
    }

    if ($self->{_pattern}  <= 0x01 and
        $self->{_bg_color} == 0x41 and
        $self->{_fg_color} != 0x40    )
    {
        $self->{_bg_color} = 0x40;
        $self->{_pattern}  = 1;
    }


    $record         = 0x00E0;
    $length         = 0x0010;

    $ifnt           = $self->{_font_index};
    $ifmt           = $self->{_num_format};


    $align          = $self->{_text_h_align};
    $align         |= $self->{_text_wrap}     << 3;
    $align         |= $self->{_text_v_align}  << 4;
    $align         |= $self->{_text_justlast} << 7;
    $align         |= $self->{_rotation}      << 8;
    $align         |= $atr_num                << 10;
    $align         |= $atr_fnt                << 11;
    $align         |= $atr_alc                << 12;
    $align         |= $atr_bdr                << 13;
    $align         |= $atr_pat                << 14;
    $align         |= $atr_prot               << 15;


    $icv            = $self->{_fg_color};
    $icv           |= $self->{_bg_color}      << 7;


    $fill           = $self->{_pattern};
    $fill          |= $self->{_bottom}        << 6;
    $fill          |= $self->{_bottom_color}  << 9;


    $border1        = $self->{_top};
    $border1       |= $self->{_left}          << 3;
    $border1       |= $self->{_right}         << 6;
    $border1       |= $self->{_top_color}     << 9;


    $border2        = $self->{_left_color};
    $border2       |= $self->{_right_color}   << 7;


    my $header      = pack("vv",       $record, $length);
    my $data        = pack("vvvvvvvv", $ifnt, $ifmt, $style, $align,
                                       $icv, $fill,
                                       $border1, $border2);

    return($header . $data);
}


###############################################################################
#
# Note to porters. The majority of the set_property() methods are created
# dynamically via Perl' AUTOLOAD sub, see below. You may prefer/have to specify
# them explicitly in other implementation languages.
#


###############################################################################
#
# get_font()
#
# Generate an Excel BIFF FONT record.
#
sub get_font {

    my $self      = shift;

    my $record;     # Record identifier
    my $length;     # Record length

    my $dyHeight;   # Height of font (1/20 of a point)
    my $grbit;      # Font attributes
    my $icv;        # Index to color palette
    my $bls;        # Bold style
    my $sss;        # Superscript/subscript
    my $uls;        # Underline
    my $bFamily;    # Font family
    my $bCharSet;   # Character set
    my $reserved;   # Reserved
    my $cch;        # Length of font name
    my $rgch;       # Font name


    $dyHeight   = $self->{_size} * 20;
    $icv        = $self->{_color};
    $bls        = $self->{_bold};
    $sss        = $self->{_font_script};
    $uls        = $self->{_underline};
    $bFamily    = $self->{_font_family};
    $bCharSet   = $self->{_font_charset};
    $rgch       = $self->{_font};

    $cch        = length($rgch);
    $record     = 0x31;
    $length     = 0x0F + $cch;
    $reserved   = 0x00;

    $grbit      = 0x00;
    $grbit     |= 0x02 if $self->{_italic};
    $grbit     |= 0x08 if $self->{_font_strikeout};
    $grbit     |= 0x10 if $self->{_font_outline};
    $grbit     |= 0x20 if $self->{_font_shadow};


    my $header  = pack("vv",         $record, $length);
    my $data    = pack("vvvvvCCCCC", $dyHeight, $grbit, $icv, $bls,
                                     $sss, $uls, $bFamily,
                                     $bCharSet, $reserved, $cch);

    return($header . $data. $self->{_font});
}

###############################################################################
#
# get_font_key()
#
# Returns a unique hash key for a font. Used by Workbook->_store_all_fonts()
#
sub get_font_key {

    my $self    = shift;

    # The following elements are arranged to increase the probability of
    # generating a unique key. Elements that hold a large range of numbers
    # eg. _color are placed between two binary elements such as _italic
    #
    my $key = "$self->{_font}$self->{_size}";
    $key   .= "$self->{_font_script}$self->{_underline}";
    $key   .= "$self->{_font_strikeout}$self->{_bold}$self->{_font_outline}";
    $key   .= "$self->{_font_family}$self->{_font_charset}";
    $key   .= "$self->{_font_shadow}$self->{_color}$self->{_italic}";
    $key    =~ s/ /_/g; # Convert the key to a single word

    return $key;
}


###############################################################################
#
# get_xf_index()
#
# Returns the used by Worksheet->_XF()
#
sub get_xf_index {
    my $self   = shift;

    return $self->{_xf_index};
}


###############################################################################
#
# _get_color()
#
# Used in conjunction with the set_xxx_color methods to convert a color
# string into a number. Color range is 0..63 but we will restrict it
# to 8..63 to comply with Gnumeric. Colors 0..7 are repeated in 8..15.
#
sub _get_color {

    my %colors = (
                    aqua    => 0x0F,
                    cyan    => 0x0F,
                    black   => 0x08,
                    blue    => 0x0C,
                    brown   => 0x10,
                    magenta => 0x0E,
                    fuchsia => 0x0E,
                    gray    => 0x17,
                    grey    => 0x17,
                    green   => 0x11,
                    lime    => 0x0B,
                    navy    => 0x12,
                    orange  => 0x35,
                    purple  => 0x14,
                    red     => 0x0A,
                    silver  => 0x16,
                    white   => 0x09,
                    yellow  => 0x0D,
                 );

    # Return the default color, 0x7FFF, if undef,
    return 0x7FFF unless defined $_[0];

    # or the color string converted to an integer,
    return $colors{lc($_[0])} if exists $colors{lc($_[0])};

    # or the default color if string is unrecognised,
    return 0x7FFF if ($_[0] =~ m/\D/);

    # or an index < 8 mapped into the correct range,
    return $_[0] + 8 if $_[0] < 8;

    # or the default color if arg is outside range,
    return 0x7FFF if $_[0] > 63;

    # or an integer in the valid range
    return $_[0];
}


###############################################################################
#
# set_align()
#
# Set cell alignment.
#
sub set_align {

    my $self     = shift;
    my $location = $_[0];

    return if not defined $location;  # No default
    return if $location =~ m/\d/;     # Ignore numbers

    $location = lc($location);

    $self->set_text_h_align(1) if ($location eq 'left');
    $self->set_text_h_align(2) if ($location eq 'centre');
    $self->set_text_h_align(2) if ($location eq 'center');
    $self->set_text_h_align(3) if ($location eq 'right');
    $self->set_text_h_align(4) if ($location eq 'fill');
    $self->set_text_h_align(5) if ($location eq 'justify');
    $self->set_text_h_align(6) if ($location eq 'merge');
    $self->set_text_h_align(7) if ($location eq 'equal_space'); # For T.K.
    $self->set_text_v_align(0) if ($location eq 'top');
    $self->set_text_v_align(1) if ($location eq 'vcentre');
    $self->set_text_v_align(1) if ($location eq 'vcenter');
    $self->set_text_v_align(2) if ($location eq 'bottom');
    $self->set_text_v_align(3) if ($location eq 'vjustify');
    $self->set_text_v_align(4) if ($location eq 'vequal_space'); # For T.K.
}


###############################################################################
#
# set_valign()
#
# Set vertical cell alignment. This is required by the set_properties() method
# to differentiate between the vertical and horizontal properties.
#
sub set_valign {

    my $self = shift;
    $self->set_align(@_);
}


###############################################################################
#
# set_merge()
#
# This is an alias for the unintuitive set_align('merge')
#
sub set_merge {

    my $self     = shift;

    $self->set_text_h_align(6);
}


###############################################################################
#
# set_bold()
#
# Bold has a range 0x64..0x3E8.
# 0x190 is normal. 0x2BC is bold. So is an excessive use of AUTOLOAD.
#
sub set_bold {

    my $self   = shift;
    my $weight = $_[0];

    $weight = 0x2BC if not defined $weight; # Bold text
    $weight = 0x2BC if $weight == 1;        # Bold text
    $weight = 0x190 if $weight == 0;        # Normal text
    $weight = 0x190 if $weight <  0x064;    # Lower bound
    $weight = 0x190 if $weight >  0x3E8;    # Upper bound

    $self->{_bold} = $weight;
}


###############################################################################
#
# set_border($style)
#
# Set cells borders to the same style
#
sub set_border {

    my $self  = shift;
    my $style = $_[0];

    $self->set_bottom($style);
    $self->set_top($style);
    $self->set_left($style);
    $self->set_right($style);
}


###############################################################################
#
# set_border_color($color)
#
# Set cells border to the same color
#
sub set_border_color {

    my $self  = shift;
    my $color = $_[0];

    $self->set_bottom_color($color);
    $self->set_top_color($color);
    $self->set_left_color($color);
    $self->set_right_color($color);
}


###############################################################################
#
# set_properties()
#
# Convert hashes of properties to method calls.
#
sub set_properties {

    my $self = shift;

    my %properties = @_; # Merge multiple hashes into one

    while (my($key, $value) = each(%properties)) {

        # Strip leading "-" from Tk style properties eg. -color => 'red'.
        $key =~ s/^-//;


        # Make sure method names are alphanumeric characters only, in case
        # tainted data is passed to the eval().
        #
        die "Unknown method: \$self->set_$key\n" if $key =~ /\W/;


        # Evaling all $values as a strings gets around the problem of some
        # numerical format strings being evaluated as numbers, for example
        # "00000" for a zip code.
        #
        if (defined $value) {
            eval "\$self->set_$key('$value')";
        }
        else {
            eval "\$self->set_$key(undef)";
        }

        die $@ if $@; # Rethrow the eval error.
    }
}


###############################################################################
#
# AUTOLOAD. Deus ex machina.
#
# Dynamically create set methods that aren't already defined.
#
sub AUTOLOAD {

    my $self = shift;

    # Ignore calls to DESTROY
    return if $AUTOLOAD =~ /::DESTROY$/;

    # Check for a valid method names, ie. "set_xxx_yyy".
    $AUTOLOAD =~ /.*::set(\w+)/ or die "Unknown method: $AUTOLOAD\n";

    # Match the attribute, ie. "_xxx_yyy".
    my $attribute = $1;

    # Check that the attribute exists
    exists $self->{$attribute}  or die "Unknown method: $AUTOLOAD\n";

    # The attribute value
    my $value;


    # There are two types of set methods: set_property() and
    # set_property_color(). When a method is AUTOLOADED we store a new anonymous
    # sub in the appropriate slot in the symbol table. The speeds up subsequent
    # calls to the same method.
    #
    no strict 'refs'; # To allow symbol table hackery

    if ($AUTOLOAD =~ /.*::set\w+color$/) {
        # For "set_property_color" methods
        $value =  _get_color($_[0]);

        *{$AUTOLOAD} = sub {
                             my $self  = shift;

                             $self->{$attribute} = _get_color($_[0]);
                           };
    }
    else {

        $value = $_[0];
        $value = 1 if not defined $value; # The default value is always 1

        *{$AUTOLOAD} = sub {
                             my $self  = shift;
                             my $value = shift;

                             $value = 1 if not defined $value;
                             $self->{$attribute} = $value;
                           };
    }


    $self->{$attribute} = $value;
}


1;


__END__


=head1 NAME

Format - A class for defining Excel formatting.

=head1 SYNOPSIS

See the documentation for Spreadsheet::WriteExcel

=head1 DESCRIPTION

This module is used in conjunction with Spreadsheet::WriteExcel.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMIII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
