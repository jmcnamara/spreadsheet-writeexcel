package Spreadsheet::WriteExcel::Formula;

###############################################################################
#
# Formula - A class for generating Excel formulas.
#
#
# Used in conjunction with Spreadsheet::WriteExcel
#
# Copyright 2000-2001, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

require Exporter;

use strict;
use Parse::RecDescent;
use Data::Dumper;
use Carp;



use vars qw($VERSION @ISA);
@ISA = qw(Exporter);

$VERSION = '0.01';

###############################################################################
#
# Class data.
#
my $parser;
my %ptg;
my %functions;


###############################################################################
#
# For debugging.
#
my $_debug = 0;


###############################################################################
#
# new()
#
# Constructor
#
sub new {

    my $class  = $_[0];

    my $self   = {
                    _byte_order    => $_[1],
                 };

    bless $self, $class;
    return $self;
}


###############################################################################
#
# _init_parser()
#
# There is a small overhead involved in generating the parser. Therefore, the
# initialisation is delayed until a formula is required. TODO: use a pre-
# compiled header.
#
sub _init_parser {

    my $self = shift;

    $self->_initialize_hashes();

    # The parsing grammar.
    # TODO:
    #       Add support for international versions of Excel
    #
    $parser = Parse::RecDescent->new(<<'EndGrammar');

        expr:           list

        # Match arg lists such as SUM(1,2, 3)
        list:           <leftop: addition ',' addition>
                        { [ $item[1], '_arg', scalar @{$item[1]} ] }

        addition:       <leftop: multiplication add_op multiplication>

        # TODO: The add_op operators don't have equal precedence.
        add_op:         add |  sub | concat
                        | eq | ne | le | ge | lt | gt   # Order is important

        add:            '+'  { 'ptgAdd'    }
        sub:            '-'  { 'ptgSub'    }
        concat:         '&'  { 'ptgConcat' }
        eq:             '='  { 'ptgEQ'     }
        ne:             '<>' { 'ptgNE'     }
        le:             '<=' { 'ptgLE'     }
        ge:             '>=' { 'ptgGE'     }
        lt:             '<'  { 'ptgLT'     }
        gt:             '>'  { 'ptgGT'     }


        multiplication: <leftop: exponention mult_op exponention>

        mult_op:        mult  | div
        mult:           '*' { 'ptgMul' }
        div:            '/' { 'ptgDiv' }

        # Right associative
        exponention:    <rightop: factor exp_op factor>

        exp_op:         '^' { 'ptgPower' }

        factor:         number       # Order is important
                        | string
                        | range
                        | true
                        | false
                        | cell
                        | function
                        | '(' expr ')'  { [$item[2], 'ptgParen'] }

        # TODO: Define a regex that can handle embedded quotes
        string:         /"[^"]+"/     #" For editors
                        { [ '_str', $item[1]] }

        # Match float or integer
        number:          /([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?/
                        { ['_num', $item[1]] }

        # Match A1, $A1, A$1 or $A$1. Highest is column 255; IV.
        cell:           /\$?[A-V][A-Z]?\$?\d+/
                        { ['_ref', $item[1]] }

        # Match A1:C5 etc.
        range:          /\$?[A-V][A-Z]?\$?\d+:\$?[A-V][A-Z]?\$?\d+/
                        { ['_rng', $item[1]] }

        # Match a function name
        function:       /[A-Z]+/ '()'
                        { ['_fnc', $item[1]] }
                        | /[A-Z]+/ '(' expr ')'
                         { [$item[3], '_fnc', $item[1]] }
                        | /[A-Z]+/ '(' list ')'
                        { [$item[3], '_fnc', $item[1]] }

        # Just watching the screen and hacking some Perl.
        true:           'TRUE'  { [ 'ptgBool', 1 ] }

        false:          'FALSE' { [ 'ptgBool', 0 ] }

EndGrammar

    print "Init_parser.\n\n" if $_debug;
}



###############################################################################
#
# parse_formula()
#
# This is the only public method. It takes a textual description of a formula
# and returns a RPN encoded byte string.
#
sub parse_formula {

    my $self = shift;

    # Initialise the parser if this is the first call
    $self->_init_parser() if not defined $parser;

    my $formula = shift @_;
    my $str;
    my $tokens;
    print $formula, "\n" if $_debug;

    # Build the parse tree for the formula
    my $parsetree =$parser->expr($formula);

    # Check if parsing worked.
    if (defined $parsetree) {
        my @tokens = $self->_reverse_tree(@$parsetree);

        # Convert parsed tokens to a byte stream
        $str = $self->_parse_tokens(@tokens);
        $tokens = join " ", @tokens;
    }
    else {
        croak("Couldn't parse formula: $formula ")
    }

    if ($_debug) {
        print join(" ", map { sprintf "%02X", $_ } unpack("C*",$str)), "\n";
        print $tokens, "\n\n";
    }

    return $str;
}


###############################################################################
#
#  _reverse_tree()
#
# This function descends recursively through the parse tree. At each level it
# swaps the order of an operator followed by an operand.
# For example, 1+2*3 would be converted in the following sequence:
#               1 + 2 * 3
#               1 + (2 * 3)
#               1 + (2 3 *)
#               1 (2 3 *) +
#               1 2 3 * +
#
sub _reverse_tree
{
    my $self = shift;

    my @tokens;
    my @expression = @_;
    my @stack;

    while (@expression) {
        my $token = shift @expression;

        # If the token is an operator swap it with the following operand
        if (    $token eq 'ptgAdd'      ||
                $token eq 'ptgSub'      ||
                $token eq 'ptgConcat'   ||
                $token eq 'ptgMul'      ||
                $token eq 'ptgDiv'      ||
                $token eq 'ptgPower'    ||
                $token eq 'ptgEQ'       ||
                $token eq 'ptgNE'       ||
                $token eq 'ptgLE'       ||
                $token eq 'ptgGE'       ||
                $token eq 'ptgLT'       ||
                $token eq 'ptgGT')
        {
            my $operand = shift @expression;
            push @stack, $operand;
        }

        push @stack, $token;
    }

    # Recurse through the parse tree
    foreach my $token (@stack) {
        if (ref($token)) {
            push @tokens, $self->_reverse_tree(@$token);
        }
        else {
            push @tokens, $token;
        }
    }

    return  @tokens;
}


###############################################################################
#
# _parse_tokens()
#
# Convert each token or token pair to its Excel 'ptg' equivalent.
#
sub _parse_tokens {

    my $self        = shift;
    my $parse_str   = '';
    my $last_type   = '';
    my $num_args    = 0;

    while (@_) {
        my $token = shift @_;

        if ($token eq '_arg') {
            $num_args = shift @_;
        }
        elsif ($token eq 'ptgBool') {
            $token = shift @_;
            $parse_str .= $self->_convert_bool($token);
        }
        elsif ($token eq '_num') {
            $token = shift @_;
            $parse_str .= $self->_convert_number($token);
        }
        elsif ($token eq '_str') {
            $token = shift @_;
            $parse_str .= $self->_convert_string($token);
        }
        elsif ($token eq '_ref') {
            $token = shift @_;
            $parse_str .= $self->_convert_reference($token);
        }
        elsif ($token eq '_rng') {
            $token = shift @_;
            $parse_str .= $self->_convert_range($token);
        }
        elsif ($token eq '_fnc') {
            $token = shift @_;
            $parse_str .= $self->_convert_function($token, $num_args);
        }
        elsif (exists $ptg{$token}) {
            $parse_str .= pack("C", $ptg{$token});
        }
        else {
            croak("Unrecognised token: $token ");
        }
    }

    return $parse_str;
}


###############################################################################
#
# _convert_bool()
#
# Convert a boolean token to ptgBool
#
sub _convert_bool {

    my $self = shift;
    my $bool = shift;

    return pack("CC", $ptg{ptgBool}, $bool);
}


###############################################################################
#
# _convert_number()
#
# Convert a number token to ptgInt or ptgNum
#
sub _convert_number {

    my $self = shift;
    my $num  = shift;

    # Integer in the range 0..2**16-1
    if (($num =~ /^\d+$/) && ($num <= 65535)) {
        return pack("Cv", $ptg{ptgInt}, $num);
    }
    else { # A float
        $num = pack("d", $num);
        $num = reverse $num if $self->{_byte_order};
        return pack("C", $ptg{ptgNum}) . $num;
    }
}


###############################################################################
#
# _convert_string()
#
# Convert a string to a ptg Str.
#
sub _convert_string {

    my $self = shift;
    my $str  = shift;

    $str =~ s/^"//;   # Remove leading  "
    $str =~ s/"$//;   # Remove trailing "
    $str =~ s/""/"/g; # Substitute Excels escaped double quote "" for "

    my $length = length($str);
    croak("String: $str greater than 255 chars ") if $length > 255;

    return pack("CC", $ptg{ptgStr}, $length) . $str;
}


###############################################################################
#
# _convert_reference
#
# Convert an Excel reference such as A1, $B2, C$3 or $D$4 to a ptgRefV.
#
sub _convert_reference {


    my $self = shift;
    my $cell = shift;

    my ($row, $col) = $self->_cell_to_packed_rowcol($cell);

    my $ptgRefV     = pack("C", $ptg{ptgRefV});

    return $ptgRefV . $row . $col;
}


###############################################################################
#
# _convert_range()
#
# Convert an Excel range such as A1:D4 to a ptgRefV.
#
sub _convert_range {

    my $self = shift;

    my $range = shift;
    my ($cell1, $cell2) = split ':', $range;
    my ($row1, $col1)   = $self->_cell_to_packed_rowcol($cell1);
    my ($row2, $col2)   = $self->_cell_to_packed_rowcol($cell2);

    my $ptgAreaV        = pack("C", $ptg{ptgArea});

    return $ptgAreaV . $row1 . $row2 . $col1. $col2;
}


###############################################################################
#
# _convert_function()
#
# Convert a function to a ptgFunc or ptgFuncVarV depending on the number of
# args that it takes.
#
sub _convert_function {

    my $self     = shift;
    my $token    = shift;
    my $num_args = shift;

    my $args = $functions{$token}[1];

    if ($args >= 0) {
        return pack("Cv", $ptg{ptgFuncV}, $functions{$token}[0]);
    }

    if ($args == -1) {
        return pack "CCv", $ptg{ptgFuncVarV}, $num_args, $functions{$token}[0];
    }
}


###############################################################################
#
# _cell_to_rowcol($cell_ref)
#
# Convert an Excel cell reference such as A1 or $B2 or C$3 or $D$4 to a zero
# indexed row and column number. Also, returns two boolean values to indicate
# if the row or column are relative references.
#
sub _cell_to_rowcol {

    my $self = shift;
    my $cell = shift;

    $cell =~ /(\$?)([A-I]?[A-Z])(\$?)(\d+)/;

    my $col_rel = $1 eq "" ? 1 : 0;
    my $col     = $2;
    my $row_rel = $3 eq "" ? 1 : 0;
    my $row     = $4;


    # Convert base26 column string to number
    my @chars  = split //, $col;
    my $expn   = 0;
    $col       = 0;

    while (@chars) {
        my $char = pop(@chars); # LS char first
        $col += (ord($char) -ord('A') +1) * (26**$expn);
        $expn++;
    }

    # Convert 1-index to zero-index
    $row--;
    $col--;

    return $row, $col, $row_rel, $col_rel;
}


###############################################################################
#
# _cell_to_packed_rowcol($row, $col, $row_rel, $col_rel)
#
# pack() row and column into the required 3 byte format.
#
sub _cell_to_packed_rowcol {

    my $self = shift;
    my $cell = shift;

    my ($row, $col, $row_rel, $col_rel) = $self->_cell_to_rowcol($cell);

    croak("Column in: $cell greater than 255 ") if $col >= 255;
    croak("Row in: $cell greater than 16384 ") if $row >= 16384;

    # Set the high bits to indicate if row or col are relative.
    $row    |= $col_rel << 14;
    $row    |= $row_rel << 15;

    $row     = pack('v', $row);
    $col     = pack('C', $col);

    return ($row, $col);
}


###############################################################################
#
# _initialize_hashes()
#
sub _initialize_hashes {

    # The Excel ptg indices
    %ptg = (
        'ptgExp'            => 0x01,
        'ptgTbl'            => 0x02,
        'ptgAdd'            => 0x03,
        'ptgSub'            => 0x04,
        'ptgMul'            => 0x05,
        'ptgDiv'            => 0x06,
        'ptgPower'          => 0x07,
        'ptgConcat'         => 0x08,
        'ptgLT'             => 0x09,
        'ptgLE'             => 0x0A,
        'ptgEQ'             => 0x0B,
        'ptgGE'             => 0x0C,
        'ptgGT'             => 0x0D,
        'ptgNE'             => 0x0E,
        'ptgIsect'          => 0x0F,
        'ptgUnion'          => 0x10,
        'ptgRange'          => 0x11,
        'ptgUplus'          => 0x12,
        'ptgUminus'         => 0x13,
        'ptgPercent'        => 0x14,
        'ptgParen'          => 0x15,
        'ptgMissArg'        => 0x16,
        'ptgStr'            => 0x17,
        'ptgAttr'           => 0x19,
        'ptgSheet'          => 0x1A,
        'ptgEndSheet'       => 0x1B,
        'ptgErr'            => 0x1C,
        'ptgBool'           => 0x1D,
        'ptgInt'            => 0x1E,
        'ptgNum'            => 0x1F,
        'ptgArray'          => 0x20,
        'ptgFunc'           => 0x21,
        'ptgFuncVar'        => 0x22,
        'ptgName'           => 0x23,
        'ptgRef'            => 0x24,
        'ptgArea'           => 0x25,
        'ptgMemArea'        => 0x26,
        'ptgMemErr'         => 0x27,
        'ptgMemNoMem'       => 0x28,
        'ptgMemFunc'        => 0x29,
        'ptgRefErr'         => 0x2A,
        'ptgAreaErr'        => 0x2B,
        'ptgRefN'           => 0x2C,
        'ptgAreaN'          => 0x2D,
        'ptgMemAreaN'       => 0x2E,
        'ptgMemNoMemN'      => 0x2F,
        'ptgNameX'          => 0x39,
        'ptgRef3d'          => 0x3A,
        'ptgArea3d'         => 0x3B,
        'ptgRefErr3d'       => 0x3C,
        'ptgAreaErr3d'      => 0x3D,
        'ptgArrayV'         => 0x40,
        'ptgFuncV'          => 0x41,
        'ptgFuncVarV'       => 0x42,
        'ptgNameV'          => 0x43,
        'ptgRefV'           => 0x44,
        'ptgAreaV'          => 0x45,
        'ptgMemAreaV'       => 0x46,
        'ptgMemErrV'        => 0x47,
        'ptgMemNoMemV'      => 0x48,
        'ptgMemFuncV'       => 0x49,
        'ptgRefErrV'        => 0x4A,
        'ptgAreaErrV'       => 0x4B,
        'ptgRefNV'          => 0x4C,
        'ptgAreaNV'         => 0x4D,
        'ptgMemAreaNV'      => 0x4E,
        'ptgMemNoMemN'      => 0x4F,
        'ptgFuncCEV'        => 0x58,
        'ptgNameXV'         => 0x59,
        'ptgRef3dV'         => 0x5A,
        'ptgArea3dV'        => 0x5B,
        'ptgRefErr3dV'      => 0x5C,
        'ptgAreaErr3d'      => 0x5D,
        'ptgArrayA'         => 0x60,
        'ptgFuncA'          => 0x61,
        'ptgFuncVarA'       => 0x62,
        'ptgNameA'          => 0x63,
        'ptgRefA'           => 0x64,
        'ptgAreaA'          => 0x65,
        'ptgMemAreaA'       => 0x66,
        'ptgMemErrA'        => 0x67,
        'ptgMemNoMemA'      => 0x68,
        'ptgMemFuncA'       => 0x69,
        'ptgRefErrA'        => 0x6A,
        'ptgAreaErrA'       => 0x6B,
        'ptgRefNA'          => 0x6C,
        'ptgAreaNA'         => 0x6D,
        'ptgMemAreaNA'      => 0x6E,
        'ptgMemNoMemN'      => 0x6F,
        'ptgFuncCEA'        => 0x78,
        'ptgNameXA'         => 0x79,
        'ptgRef3dA'         => 0x7A,
        'ptgArea3dA'        => 0x7B,
        'ptgRefErr3dA'      => 0x7C,
        'ptgAreaErr3d'      => 0x7D,
    );

    # The Excel function names with their index code and a value to indicate
    # the number of arguments that the take:
    #    >=0 is a fixed number of arguments
    #    -1  is variable
    #
    # Thanks to Michael Meeks and Gnumeric for the number of arg values.
    #
    %functions  = (
        'COUNT'             => [   0,  -1 ],
        'IF'                => [   1,  -1 ],
        'ISNA'              => [   2,   1 ],
        'ISERROR'           => [   3,   1 ],
        'SUM'               => [   4,  -1 ],
        'AVERAGE'           => [   5,  -1 ],
        'MIN'               => [   6,  -1 ],
        'MAX'               => [   7,  -1 ],
        'ROW'               => [   8,   1 ],
        'COLUMN'            => [   9,   1 ],
        'NA'                => [  10,   0 ],
        'NPV'               => [  11,  -1 ],
        'STDEV'             => [  12,  -1 ],
        'DOLLAR'            => [  13,  -1 ],
        'FIXED'             => [  14,  -1 ],
        'SIN'               => [  15,   1 ],
        'COS'               => [  16,   1 ],
        'TAN'               => [  17,   1 ],
        'ATAN'              => [  18,   1 ],
        'PI'                => [  19,   0 ],
        'SQRT'              => [  20,   1 ],
        'EXP'               => [  21,   1 ],
        'LN'                => [  22,   1 ],
        'LOG10'             => [  23,   1 ],
        'ABS'               => [  24,   1 ],
        'INT'               => [  25,   1 ],
        'SIGN'              => [  26,   1 ],
        'ROUND'             => [  27,   2 ],
        'LOOKUP'            => [  28,  -1 ],
        'INDEX'             => [  29,  -1 ],
        'REPT'              => [  30,   2 ],
        'MID'               => [  31,   3 ],
        'LEN'               => [  32,   1 ],
        'VALUE'             => [  33,   1 ],
        'TRUE'              => [  34,   0 ],
        'FALSE'             => [  35,   0 ],
        'AND'               => [  36,  -1 ],
        'OR'                => [  37,  -1 ],
        'NOT'               => [  38,   1 ],
        'MOD'               => [  39,   2 ],
        'DCOUNT'            => [  40,   3 ],
        'DSUM'              => [  41,   3 ],
        'DAVERAGE'          => [  42,   3 ],
        'DMIN'              => [  43,   3 ],
        'DMAX'              => [  44,   3 ],
        'DSTDEV'            => [  45,   3 ],
        'VAR'               => [  46,  -1 ],
        'DVAR'              => [  47,   3 ],
        'TEXT'              => [  48,   2 ],
        'LINEST'            => [  49,  -1 ],
        'TREND'             => [  50,  -1 ],
        'LOGEST'            => [  51,  -1 ],
        'GROWTH'            => [  52,  -1 ],
        'PV'                => [  56,  -1 ],
        'FV'                => [  57,  -1 ],
        'NPER'              => [  58,  -1 ],
        'PMT'               => [  59,  -1 ],
        'RATE'              => [  60,  -1 ],
        'MIRR'              => [  61,   3 ],
        'IRR'               => [  62,  -1 ],
        'RAND'              => [  63,   0 ],
        'MATCH'             => [  64,   3 ],
        'DATE'              => [  65,   3 ],
        'TIME'              => [  66,   3 ],
        'DAY'               => [  67,   1 ],
        'MONTH'             => [  68,   1 ],
        'YEAR'              => [  69,   1 ],
        'WEEKDAY'           => [  70,  -1 ],
        'HOUR'              => [  71,   1 ],
        'MINUTE'            => [  72,   1 ],
        'SECOND'            => [  73,   1 ],
        'NOW'               => [  74,   0 ],
        'AREAS'             => [  75,   1 ],
        'ROWS'              => [  76,   1 ],
        'COLUMNS'           => [  77,   1 ],
        'OFFSET'            => [  78,  -1 ],
        'SEARCH'            => [  82,  -1 ],
        'TRANSPOSE'         => [  83,   1 ],
        'TYPE'              => [  86,   1 ],
        'ATAN2'             => [  97,   2 ],
        'ASIN'              => [  98,   1 ],
        'ACOS'              => [  99,   1 ],
        'CHOOSE'            => [ 100,  -1 ],
        'HLOOKUP'           => [ 101,  -1 ],
        'VLOOKUP'           => [ 102,  -1 ],
        'ISREF'             => [ 105,   1 ],
        'LOG'               => [ 109,  -1 ],
        'CHAR'              => [ 111,   1 ],
        'LOWER'             => [ 112,   1 ],
        'UPPER'             => [ 113,   1 ],
        'PROPER'            => [ 114,   1 ],
        'LEFT'              => [ 115,  -1 ],
        'RIGHT'             => [ 116,  -1 ],
        'EXACT'             => [ 117,   2 ],
        'TRIM'              => [ 118,   1 ],
        'REPLACE'           => [ 119,   4 ],
        'SUBSTITUTE'        => [ 120,  -1 ],
        'CODE'              => [ 121,   1 ],
        'FIND'              => [ 124,  -1 ],
        'CELL'              => [ 125,   2 ],
        'ISERR'             => [ 126,   1 ],
        'ISTEXT'            => [ 127,   1 ],
        'ISNUMBER'          => [ 128,   1 ],
        'ISBLANK'           => [ 129,   1 ],
        'T'                 => [ 130,   1 ],
        'N'                 => [ 131,   1 ],
        'DATEVALUE'         => [ 140,   1 ],
        'TIMEVALUE'         => [ 141,   1 ],
        'SLN'               => [ 142,   3 ],
        'SYD'               => [ 143,   4 ],
        'DDB'               => [ 144,  -1 ],
        'INDIRECT'          => [ 148,  -1 ],
        'CALL'              => [ 150,  -1 ],
        'CLEAN'             => [ 162,   1 ],
        'MDETERM'           => [ 163,   1 ],
        'MINVERSE'          => [ 164,   1 ],
        'MMULT'             => [ 165,   2 ],
        'IPMT'              => [ 167,  -1 ],
        'PPMT'              => [ 168,  -1 ],
        'COUNTA'            => [ 169,  -1 ],
        'PRODUCT'           => [ 183,  -1 ],
        'FACT'              => [ 184,   1 ],
        'DPRODUCT'          => [ 189,   3 ],
        'ISNONTEXT'         => [ 190,   1 ],
        'STDEVP'            => [ 193,  -1 ],
        'VARP'              => [ 194,  -1 ],
        'DSTDEVP'           => [ 195,   3 ],
        'DVARP'             => [ 196,   3 ],
        'TRUNC'             => [ 197,  -1 ],
        'ISLOGICAL'         => [ 198,   1 ],
        'DCOUNTA'           => [ 199,   3 ],
        'ROUNDUP'           => [ 212,   2 ],
        'ROUNDDOWN'         => [ 213,   2 ],
        'RANK'              => [ 216,  -1 ],
        'ADDRESS'           => [ 219,  -1 ],
        'DAYS360'           => [ 220,  -1 ],
        'TODAY'             => [ 221,   0 ],
        'VDB'               => [ 222,  -1 ],
        'MEDIAN'            => [ 227,  -1 ],
        'SUMPRODUCT'        => [ 228,  -1 ],
        'SINH'              => [ 229,   1 ],
        'COSH'              => [ 230,   1 ],
        'TANH'              => [ 231,   1 ],
        'ASINH'             => [ 232,   1 ],
        'ACOSH'             => [ 233,   1 ],
        'ATANH'             => [ 234,   1 ],
        'DGET'              => [ 235,   3 ],
        'INFO'              => [ 244,   1 ],
        'DB'                => [ 247,  -1 ],
        'FREQUENCY'         => [ 252,   2 ],
        'ERROR.TYPE'        => [ 261,   1 ],
        'REGISTER.ID'       => [ 267,  -1 ],
        'AVEDEV'            => [ 269,  -1 ],
        'BETADIST'          => [ 270,  -1 ],
        'GAMMALN'           => [ 271,   1 ],
        'BETAINV'           => [ 272,  -1 ],
        'BINOMDIST'         => [ 273,   4 ],
        'CHIDIST'           => [ 274,   2 ],
        'CHIINV'            => [ 275,   2 ],
        'COMBIN'            => [ 276,   2 ],
        'CONFIDENCE'        => [ 277,   3 ],
        'CRITBINOM'         => [ 278,   3 ],
        'EVEN'              => [ 279,   1 ],
        'EXPONDIST'         => [ 280,   3 ],
        'FDIST'             => [ 281,   3 ],
        'FINV'              => [ 282,   3 ],
        'FISHER'            => [ 283,   1 ],
        'FISHERINV'         => [ 284,   1 ],
        'FLOOR'             => [ 285,   2 ],
        'GAMMADIST'         => [ 286,   4 ],
        'GAMMAINV'          => [ 287,   3 ],
        'CEILING'           => [ 288,   2 ],
        'HYPGEOMDIST'       => [ 289,   4 ],
        'LOGNORMDIST'       => [ 290,   3 ],
        'LOGINV'            => [ 291,   3 ],
        'NEGBINOMDIST'      => [ 292,   3 ],
        'NORMDIST'          => [ 293,   4 ],
        'NORMSDIST'         => [ 294,   1 ],
        'NORMINV'           => [ 295,   3 ],
        'NORMSINV'          => [ 296,   1 ],
        'STANDARDIZE'       => [ 297,   3 ],
        'ODD'               => [ 298,   1 ],
        'PERMUT'            => [ 299,   2 ],
        'POISSON'           => [ 300,   3 ],
        'TDIST'             => [ 301,   3 ],
        'WEIBULL'           => [ 302,   4 ],
        'SUMXMY2'           => [ 303,   2 ],
        'SUMX2MY2'          => [ 304,   2 ],
        'SUMX2PY2'          => [ 305,   2 ],
        'CHITEST'           => [ 306,   2 ],
        'CORREL'            => [ 307,   2 ],
        'COVAR'             => [ 308,   2 ],
        'FORECAST'          => [ 309,   3 ],
        'FTEST'             => [ 310,   2 ],
        'INTERCEPT'         => [ 311,   2 ],
        'PEARSON'           => [ 312,   2 ],
        'RSQ'               => [ 313,   2 ],
        'STEYX'             => [ 314,   2 ],
        'SLOPE'             => [ 315,   2 ],
        'TTEST'             => [ 316,   4 ],
        'PROB'              => [ 317,  -1 ],
        'DEVSQ'             => [ 318,  -1 ],
        'GEOMEAN'           => [ 319,  -1 ],
        'HARMEAN'           => [ 320,  -1 ],
        'SUMSQ'             => [ 321,  -1 ],
        'KURT'              => [ 322,  -1 ],
        'SKEW'              => [ 323,  -1 ],
        'ZTEST'             => [ 324,  -1 ],
        'LARGE'             => [ 325,   2 ],
        'SMALL'             => [ 326,   2 ],
        'QUARTILE'          => [ 327,   2 ],
        'PERCENTILE'        => [ 328,   2 ],
        'PERCENTRANK'       => [ 329,  -1 ],
        'MODE'              => [ 330,  -1 ],
        'TRIMMEAN'          => [ 331,   2 ],
        'TINV'              => [ 332,   2 ],
        'CONCATENATE'       => [ 336,  -1 ],
        'POWER'             => [ 337,   2 ],
        'RADIANS'           => [ 342,   1 ],
        'DEGREES'           => [ 343,   1 ],
        'SUBTOTAL'          => [ 344,  -1 ],
        'SUMIF'             => [ 345,  -1 ],
        'COUNTIF'           => [ 346,   2 ],
        'COUNTBLANK'        => [ 347,   1 ],
        'ROMAN'             => [ 354,  -1 ],
    );

}




1;


__END__


=head1 NAME

Formula - A class for generating Excel formulas

=head1 SYNOPSIS

See the documentation for Spreadsheet::WriteExcel

=head1 DESCRIPTION

This module is used in conjunction with Spreadsheet::WriteExcel. It should not be used directly.

The following notes are to help developers and maintainers understand the sequence of operation. They are also intended as a pro-memoria for the author. ;-)

Spreadsheet::WriteExcel::Formula converts a textual representation of a formula into the pre-parsed binary format that Excel uses to store formulas. For example C<1+2*3> is stored as follows: C<1E 01 00 1E 02 00 1E 03 00 05 03>. This string is comprised of operators and operands arranged in a reverse-Polish format. The meaning of the tokens in the above example is shown in the following table:

    Token   Name        Value
    1E      ptgInt      0001   (stored as 01 00)
    1E      ptgInt      0002   (stored as 02 00)
    1E      ptgInt      0003   (stored as 03 00)
    05      ptgMul
    03      ptgAdd

The tokens and token names are defined in the "Excel Developer's Kit" from Microsoft Press. C<ptg> stands for Parse ThinG (as in "That lexer can't grok it, it's a parse thang.")

In general the tokens fall into two categories: operators such as C<ptgMul> above and operands such as C<ptgInt>. When the formula is evaluated by Excel the operand tokens push values onto a stack. The operator tokens then pop the required number of operands from the stack, perform an operation and push the resulting value back onto the stack. This methodology is similar to the basic operation of a reverse-Polish (RPN) calculator.

Spreadsheet::WriteExcel::Formula parses a formula using a C<Parse::RecDescent> parser (at a later stage it may use a C<Parse::Yapp> parser).

The parser converts the textual representation of a formula into a parse tree. Thus, C<1+2*3> is converted into something like the following, C<e> stands for expression:

             e
           / | \
         1   +   e
               / | \
             2   *   3


The function C<_reverse_tree()> recurses down through this structure swapping the order of operators followed by operands to produce a reverse-Polish tree. Following the above example the resulting tree would look like this:


             e
           / | \
         1   e   +
           / | \
         2   3   *

The result of the recursion is a single array of tokens. In our example the simplified form would look like the following:

    (1, 2, 3, *, +)

The actual return value contains some additional information to help in the secondary parsing stage:

    (_num, 1, _num, 2, _num, 3, ptgMul, ptgAdd, _arg, 1)

The additional tokens are:

    Token   Meaning

    _num    The next token is a number
    _str    The next token is a string
    _ref    The next token is a cell reference
    _rng    The next token is a range
    _fnc    The next token is a function
    _arg    The next token is a the number of args for a function

The C<_arg> token is generated for all lists but is only used for functions the take a variable number of arguments.


A secondary parsing stage is carried out by C<_parse_tokens()> which converts these tokens into a binary string. For the C<1+2*3> example this would give:

    1E 01 00 1E 02 00 1E 03 00 05 03

This two-pass method could probably have been reduced to a single pass through the C<Parse::RecDescent> parser. However, it was easier to develop and debug this way.

The token values and formula values are stored in the C<%ptg> and C<%functions> hashes. These hashes and the parser object C<$parser> are exposed as global data. This breaks the OO encapsulation, but means that they can be shared by several instances of Spreadsheet::WriteExcel called from the same program.

The parser is initialised by C<_init_parser()>. The initialisation is delayed until the first formula is parsed. This eliminates the overhead of generating the parser in programs that are not processing formulas. (The parser should really be pre-compiled, this is to-do when the grammar stabilises).


=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMI, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
