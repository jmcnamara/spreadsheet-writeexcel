package Spreadsheet::WriteExcel;

######################################################################
#
# WriteExcel.
#
# Spreadsheet::WriteExcel - Write formatted text and numbers to a
# cross-platform Excel binary file.
#
# Copyright 2000, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

require Exporter;

use strict;
use Spreadsheet::Workbook;




use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::Workbook Exporter);

$VERSION = '0.22'; # 22 October 2000, Roethke

######################################################################
#
# new()
#
# Constructor. Wrapper for a Workbook object.
# uses: Spreadsheet::BIFFwriter
#       Spreadsheet::OLEwriter
#       Spreadsheet::Workbook
#       Spreadsheet::Worksheet
#       Spreadsheet::Format
#
sub new {

    my $class = shift;
    my $self  = Spreadsheet::Workbook->new($_[0]);

    bless  $self, $class;
    return $self;
}


1;


__END__



=head1 NAME

Spreadsheet::WriteExcel - Write formatted text and numbers to a cross-platform Excel binary file.




=head1 VERSION

This document refers to version 0.22 of Spreadsheet::WriteExcel, released October 22, 2000.




=head1 SYNOPSIS

To write a string, a number and a formatted string to the first worksheet in an Excel workbook called perl.xls:

    use Spreadsheet::WriteExcel;

    $row1 = $col1 = 0;
    $row2 = 1;
    $row3 = 2;

    my $workbook = Spreadsheet::WriteExcel->new("perl.xls");
    $worksheet   = $workbook->addworksheet();
    $format      = $workbook->addformat();
    
    $format->set_bold();
    $format->set_color('red');
    $format->set_align('center');

    $worksheet->write($row1, $col1, "Hi Excel!");
    $worksheet->write($row2, $col1, 1.2345);
    $worksheet->write($row3, $col1, "Hi Excel!", $format);


=head1 DESCRIPTION

The Spreadsheet::WriteExcel module can be used to write numbers and text in the native Excel binary file format. Multiple worksheets can be added to a workbook and formatting can be applied to cells.

The Excel file produced by this module is compatible with Excel 5, 95, 97 and 2000.

The module will work on the majority of Windows, UNIX and Macintosh platforms. Generated files are also compatible with the Linux/UNIX spreadsheet applications Star Office, Gnumeric and XESS. The generated files are not compatible with MS Access.

An article describing Spreadsheet::WriteExcel appears in Issue #19 of The Perl Journal, http://www.itknowledge.com/tpj/


=head1 WORKBOOK METHODS

The Spreadsheet::WriteExcel module provides an object oriented interface to a new Excel workbook.The following methods are available through a new workbook.

If you are unfamiliar with object oriented interfaces or the way that they are implemented in Perl have a look at C<perlobj> and C<perltoot> in the main Perl documentation.


=head2 new()

A new Excel workbook is created using the C<new()> constructor as follows:

    my $workbook = Spreadsheet::WriteExcel->new("filename.xls");

Note C<my> is required to allocate a new workbook regardless of whether the C<strict> pragma is in operation or not.

You can  redirect the output to STDOUT using the special Perl filehandle C<"-">. This can be useful for CGIs which have a Content-type of C<application/vnd.ms-excel>, for example:

    #!/usr/bin/perl -w

    use strict;
    use Spreadsheet::WriteExcel;

    print "Content-type: application/vnd.ms-excel\n\n";

    my $workbook = Spreadsheet::WriteExcel->new("-");
    $workbook->write(0, 0, "Hi Excel!");




=head2 addworksheet($sheetname)

At least one worksheet should be added to a new workbook:

    $worksheet1 = $workbook->addworksheet();          # Sheet1
    $worksheet2 = $workbook->addworksheet('Foglio2'); # Foglio2
    $worksheet3 = $workbook->addworksheet('Data');    # Data
    $worksheet4 = $workbook->addworksheet();          # Sheet4

If C<$sheetname> is not specified the default Excel convention will be followed, i.e. Sheet1, Sheet2, etc.




=head2 addformat()

The C<addformat()> method can be used to create new Format objects which are used to apply formatting to a cell:

    $format1 = $workbook->addformat();
    $format2 = $workbook->addformat();

See the L<FORMAT METHODS> section for details.




=head2 close()

The C<close()> method can be called to explicitly close an Excel file. Otherwise the file will be closed automatically when the object reference goes out of scope or when the program ends.

    $workbook->close();

In general it is only necessary to explicitly close a file if you want to perform some other operation on it such as copying or checking the size. HOWEVER, there is a bug in the current implementation that sometimes causes a file of 0 bytes to be written. This can be avoided by explicitly calling the C<close()> method for the WriteExcel object. See the L<BUGS> section.




=head2 worksheets()

The C<worksheets()> method returns a reference to the array of worksheets in a workbook. This can be useful if you want to repeat an operation on each worksheet in a workbook or if you wish to refer to a worksheet by its index:

    foreach $worksheet (@{$workbook->worksheets()}) {
       $worksheet->write(0, 0, "Hello");
    }
    
    # or:
    
    $worksheets = $workbook->worksheets();
    @$worksheets[0]->write(0, 0, "Hello");


References are explained in detail in C<perlref> and C<perlreftut> in the main Perl documentation.




=head1 WORKSHEET METHODS

The following methods are available through a new worksheet. A new worksheet is created by calling the C<addworksheet()> method from a Workbook object:

    $worksheet1 = $workbook->addworksheet();
    $worksheet2 = $workbook->addworksheet();





=head2 write($row, $column, $token, $format)

The  C<write()> method calls C<write_number()> if C<$token> matches the following regex:

    $token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/

Otherwise it calls C<write_string()>:

    $worksheet->write(0, 0, "Hello" );  # write_string()
    $worksheet->write(1, 0, "One"   );  # write_string()
    $worksheet->write(2, 0,  2      );  # write_number()
    $worksheet->write(3, 0,  3.00001);  # write_number()

The C<$format> argument is optional. It should be a valid Format object, see L<FORMAT METHODS>:

    my $format = $workbook->addformat();
    $format->set_bold();
    $format->set_color('red');
    $format->set_align('center');

    $worksheet->write(4, 0, "Hello", $format ); # Formatted string


It should be noted that C<$row> and C<$column> are zero indexed cell locations for the C<write> methods. Thus, Cell A1 is (0, 0) and Cell AD2000 is (1999, 29). Cells can be written to in any order but for forward compatibility it is probably best to write them in row-column order when possible.

The C<write> methods return:

    0 for success
   -1 for insufficient number of arguments
   -2 for row or column out of bounds
   -3 for string too long.




=head2 write_number($row, $column, $number, $format)

Write an integer or a float to the cell specified by C<$row> and C<$column>:

    $worksheet->write_number(0, 0,  1     );
    $worksheet->write_number(1, 0,  2.3451);

The C<$format> argument is optional.




=head2 write_string($row, $column, $string, $format)

Write a string to the cell specified by C<$row> and C<$column>:

    $worksheet->write_string(0, 0, "Your text here" );

The maximum string size is 255 characters. The C<$format> argument is optional.




=head2 write_blank($row, $column, $format)

Write a blank cell specified by C<$row> and C<$column>:

    $worksheet->write_blank(0, 0, $format);

This method is useful for adding formatting  to a cell that doesn't contain a string or number value.




=head2 activate()

The C<activate()> method is used to specify which worksheet is initially selected in a multi-sheet workbook:

    $worksheet1 = $workbook->addworksheet('To');
    $worksheet2 = $workbook->addworksheet('the');
    $worksheet3 = $workbook->addworksheet('wind');

    $worksheet3->activate();

This is similar to the Excel VBA activate method. The default value is the first worksheet.




=head2 set_first_sheet()

The C<activate()> method determines which worksheet is initially selected. However, if there are a large number of worksheets the selected worksheet may not appear on the screen. To avoid this you can select which is the leftmost visible worksheet using C<set_first_sheet()>:

    for (1..20) {
        $workbook->addworksheet;
    }

    $worksheet21 = $workbook->addworksheet();
    $worksheet22 = $workbook->addworksheet();

    $worksheet21->set_first_sheet();
    $worksheet22->activate();

This method is not required very often. The default value is the first worksheet.




=head2 set_selection($first_row, $first_col, $last_row, $last_col);

This method can be used to specify which cell or cells are selected in a worksheet. The most common requirement is to select a single cell, in which case C<$last_row> and C<$last_col> are not required. The active cell within a selected range is determined by the order in which C<$first> and C<$last> are specified:

    $worksheet1->set_selection(3, 3);
    $worksheet2->set_selection(3, 3, 6, 6);
    $worksheet3->set_selection(6, 6, 3, 3);

The default is cell (0, 0).




=head2 set_col_width($first_col, $last_col, $width, $format);

This method can be used to specify the width of a single column or a range of columns. If the method is applied to a single column the value of C<$first_col> and C<$last_col> should be the same:

    $worksheet->set_col_width(0, 0, 20);
    $worksheet->set_col_width(1, 3, 30);

The width corresponds to the column width value that is specified in Excel. It is approximately equal to the length of a string in the default font of Arial 10. The C<$format> argument is optional, for additional information, see L<FORMAT METHODS>.





=head1 FORMAT METHODS

This section describes the methods that are available through a Format object. Format objects are created by calling the Workbook C<addformat()> method as follows:

    my $heading1 = $workbook->addformat();
    my $heading2 = $workbook->addformat();

The format object holds all the formatting properties that can be applied to a cell, a row or a column. The following table shows the Excel format categories, the formatting properties that can be applied and the relevant object method to do so:


    Category        Property            Method Name
    --------        --------            -----------
    Font            Font type           set_font()
                    Font size           set_size()
                    Font color          set_color()
                    Bold                set_bold()
                    Italic              set_italic()
                    Underline           set_underline()
                    Strikeout           set_font_strikeout()
                    Super/Subscript     set_font_script()
                    Outline             set_font_outline()
                    Shadow              set_font_shadow()

    Number          Numeric format      set_format()
    
    Alignment       Horizontal align    set_align()
                    Vertical align      set_align()
                    Rotation            set_rotation()
                    Text wrap           set_text_wrap()
                    Justify last        set_text_justlast()

    Pattern         Cell pattern        set_pattern()
                    Background color    set_bg_color()
                    Foreground color    set_fg_color()
    
    Border          Cell border         set_border()
                    Bottom border       set_bottom()
                    Top border          set_top()
                    Left border         set_left()
                    Right border        set_right()
                    Border color        set_border_color()
                    Bottom color        set_bottom_color()
                    Top color           set_top_color()
                    Left color          set_left_color()
                    Right color         set_right_color()


The default format is Arial 10 with all other properties off. In general a method call without an argument will turn a property on, for example:

    my $format1 = $workbook->addformat();
    $format1->set_bold();  # Turns bold on
    $format1->set_bold(1); # Also turns bold on
    $format1->set_bold(0); # Turns bold off

More than one property can be applied to a format:

    my $format2 = $workbook->addformat();
    $format2->set_bold();
    $format2->set_italic();
    $format2->set_color('red');

Once a Format object has been constructed it can be passed as an argument to the Worksheet C<write> methods as follows:

    $worksheet->write(0, 0, "One", $format);
    $worksheet->write_string(1, 0, "Two", $format);
    $worksheet->write_number(2, 0, 3, $format);
    $worksheet->write_blank(3, 0, $format);

Formats can also be passed to the Worksheet C<set_col_width()> method to define the default property for a column.

    $worksheet->set_col_width(0, 0, 15, $format);

However, this will not currently set the property for cells in the column written by WriteExcel, it only applies to cells written after the workbook is opened in Excel.

    $worksheet->set_col_width(0, 0, 15, $format);

NOTE: It is important to understand that a Format is applied to a cell not in its current state  but in its final state. Consider the following example:

    my $format = $workbook->addformat();
    $format->set_bold();
    $format->set_color('red');
    $worksheet->write(0, 0, "Cell A1", $format);
    $format->set_color('green');
    $worksheet->write(0, 1, "Cell B1", $format);

Cell A1 is assigned the Format C<$format> which is initially set to the colour red. However, the colour is subsequently set to green. When Excel displays Cell A1 it will display the final state of the Format which in this case will be the colour green.

The Format object methods are described in more detail in the following sections. In addition, there is a Perl program in the WriteExcel distribution called C<formats.pl>. If you run this program it creates an Excel workbook called C<formats.xls> that contains examples of all possible format types.




=head2 copy($format)


This is the only method of a Format object that doesn't apply directly to a property. It is used to copy all of the properties from one Format object to another:

    my $lorry1 = $workbook->addformat();
    $lorry1->set_bold();
    $lorry1->set_italic();
    $lorry1->set_color('red');    # lorry1 is bold, italic and red

    my $lorry2 = $workbook->addformat();
    $lorry2->copy($lorry1);
    $lorry2->set_color('yellow'); # lorry2 is bold, italic and yellow

This can be useful when you are setting up several complex but similar formats. Note, this is not a copy constructor, both objects must exist prior to copying.




=head2 set_font($fontname)

    Default state:      Font is Arial
    Default action:     None
    Valid args:         Any valid font name


Excel can only display fonts that are installed on the system that it is running on. Therefore it is best to stick to the fonts that come as standard such as 'Arial', 'Times New Roman' and 'Courier New'. For examples see the Fonts worksheet created by formats.pl




=head2 set_size()

    Default state:      Font size is 10
    Default action:     Set font size to 1
    Valid args:         Integer values from 1 to as big as your screen.


Excel adjusts the height of a row to accommodate the largest font size in the row. Since it is not currently possible to set the row height directly you can use this as a workaround to set the row height as follows:

    # This is a kludge to set the row height
    my $format = $workbook->addformat();
    $format->set_size(30);
    $worksheet->write_blank($row, $some_unused_cell, $format);




=head2 set_color()

    Default state:      Excels default color, usually black
    Default action:     Set the default color
    Valid args:         Integers form 6..63 or the following strings:
                        'aqua'
                        'black'
                        'blue'
                        'fuchsia'
                        'gray'
                        'green'
                        'lime'
                        'navy'
                        'orange'
                        'purple'
                        'red'
                        'silver'
                        'white'
                        'yellow'

The C<set_color()> method is used to set the colour of the font in a cell. To set the colour of a cell use the C<set_fg_color()> method.

    my $format = $workbook->addformat();
    $format->set_color('red');
    $worksheet->write(0, 0, "wheelbarrow", $format);

For additional examples see the 'Named colors' and 'Standard colors' worksheets created by formats.pl




=head2 set_bold()

    Default state:      bold is off
    Default action:     Turn bold on
    Valid args:         0, 1*

* Actually values in the range 100..1000 are also valid. 400 is normal, 700 is bold and 1000 is very bold indeed. It is probably best to set the value to 1 and use normal bold.




=head2 set_italic()

    Default state:      Italic is off
    Default action:     Turn italic on
    Valid args:         0, 1




=head2 set_underline()

    Default state:      Underline is off
    Default action:     Turn on single underline
    Valid args:         0  = No underline
                        1  = Single underline
                        2  = Double underline
                        33 = Single accounting underline
                        34 = Double accounting underline




=head2 set_strikeout()

    Default state:      Strikeout is off
    Default action:     Turn strikeout on
    Valid args:         0, 1





=head2 set_script()

    Default state:      Super/Subscript is off
    Default action:     Turn Superscript on
    Valid args:         0  = Normal
                        1  = Superscript
                        2  = Subscript

This format will not be very useful until multiple formats can be applied to a string.


=head2 set_outline()

    Default state:      Outline is off
    Default action:     Turn outline on
    Valid args:         0, 1

Macintosh only.




=head2 set_shadow()

    Default state:      Shadow is off
    Default action:     Turn shadow on
    Valid args:         0, 1

Macintosh only.




=head2 set_format($format_code)

    Default state:      General format
    Default action:     Format index 1
    Valid args:         See the following table

This method is used to define the way that a number is displayed in Excel. For example C<36831.0153> can be displayed as an integer, a floating point number, a date or a currency value. User defined numerical formats are not supported in this version. Only the built-in formats are available as shown in the following table:

    Index   Index   Format String
    0       0x00    General
    1       0x01    0
    2       0x02    0.00
    3       0x03    #,##0
    4       0x04    #,##0.00
    5       0x05    ($#,##0_);($#,##0)
    6       0x06    ($#,##0_);[Red]($#,##0)
    7       0x07    ($#,##0.00_);($#,##0.00)
    8       0x08    ($#,##0.00_);[Red]($#,##0.00)
    9       0x09    0%
    10      0x0a    0.00%
    11      0x0b    0.00E+00
    12      0x0c    # ?/?
    13      0x0d    # ??/??
    14      0x0e    m/d/yy
    15      0x0f    d-mmm-yy
    16      0x10    d-mmm
    17      0x11    mmm-yy
    18      0x12    h:mm AM/PM
    19      0x13    h:mm:ss AM/PM
    20      0x14    h:mm
    21      0x15    h:mm:ss
    22      0x16    m/d/yy h:mm
    ..      ....    ...........
    37      0x25    (#,##0_);(#,##0)
    38      0x26    (#,##0_);[Red](#,##0)
    39      0x27    (#,##0.00_);(#,##0.00)
    40      0x28    (#,##0.00_);[Red](#,##0.00)
    41      0x29    _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
    42      0x2a    _($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)
    43      0x2b    _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
    44      0x2c    _($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)
    45      0x2d    mm:ss
    46      0x2e    [h]:mm:ss
    47      0x2f    mm:ss.0
    48      0x30    ##0.0E+0
    49      0x31    @

In Excel 5 the dollar sign appears as a dollar sign. In Excel 97-2000 it appears as the defined local currency symbol.

Also, the red negative numeric formats display slightly differently in Excel 5 and Excel 97-2000.

Dates in this implementation of the Excel file are in the 1900 format. That is to say, the integer part of the number stores the number of days since 1900 while the fractional part stores the percentage of the day. If someone writes a robust and documented program for conversion between UNIX/Perl time and the 1900 system please send it along for inclusion in the WriteExcel distro.

Numeric formats 23 to 36 are not documented by Microsoft and may differ in international versions. Try them and see what you get.


The C<set_format()> method is used as follows:

    my $format = $workbook->addformat();
    $format->set_format(0x0f);
    $worksheet->write(0, 0, 36831.0153, $format); # Displays 01-Nov-00

For examples see the 'Numerical formats' worksheet created by formats.pl.




=head2 set_align()

    Default state:      Alignment is off
    Default action:     Left alignment
    Valid args:         'left'              Horizontal
                        'center'
                        'right'
                        'fill'
                        'justify'
                        'merge'
                    
                        'top'               Vertical
                        'vcenter'
                        'bottom'
                        'vjustify'

This method is used to set the horizontal and vertical text alignment within a cell. Vertical and horizontal alignments can be combined. The method is used as follows:

    my $format = $workbook->addformat();
    $format->set_align('center');
    $format->set_align('vcenter');
    $worksheet->write(0, 0, "X", $format);

Text can be aligned across two or more adjacent cells using the C<merge> property. Only one cell should contain the text, the other cells should be blank:

    my $format = $workbook->addformat();
    $format->set_align('merge');

    $worksheet->write(1, 1, 'Merged cells', $format);
    $worksheet->write_blank(1, 2, $format);

For further examples see the 'Alignment' worksheet created by formats.pl.




=head2 set_rotation()

    Default state:      Text rotation is off
    Default action:     Rotation style 1
    Valid args:         0 No rotation
                        1 Letters run from top to bottom
                        2 90° anticlockwise
                        3 90° clockwise


See the 'Alignment' worksheet created by formats.pl.




=head2 set_text_wrap()

    Default state:      Text wrap is off
    Default action:     Turn text wrap on
    Valid args:         0, 1


Here is an example using the text wrap property, the escape character C<\n> is used to indicate the end of line:

    my $format = $workbook->addformat();
    $format->set_text_wrap();
    $worksheet->write(0, 0, "It's\na bum\nwrap", $format);

Excel will adjusts the height of the row to accommodate the wrapped text.


=head2 set_text_justlast()

    Default state:      Justify last is off
    Default action:     Turn Justify last on
    Valid args:         0, 1


Only applies to Far Eastern versions of Excel.




=head2 set_pattern()

    Default state:      Pattern is off
    Default action:     Solid fill is on
    Valid args:         0 .. 31


Examples of the available patterns are shown in the 'Patterns' worksheet created by formats.pl. However, it is unlikely that you will ever need anything other than Pattern 1 which is a solid fill of the foreground color.




=head2 set_fg_color()

    Also applies to:    set_bg_color
    
    Default state:      Color is off
    Default action:     Undefined
    Valid args:         See set_color()


Here is an example of how to set the color in a cell:

    my $format = $workbook->addformat();
    $format->set_pattern(0x1);
    $format->set_fg_color('green');
    $worksheet->write(0, 0, "Ray", $format);




=head2 set_border()

    Also applies to:    set_bottom()
                        set_top()
                        set_left()
                        set_right()
                        
    Default state:      Border is off
    Default action:     Set border type 1
    Valid args:         0 No border
                        1 Thin single border
                        2 Medium single border
                        3 Dashed border
                        4 Dotted border
                        5 Thick single border
                        6 Double line border
                        7 Hair border


A cell comprises a border on the bottom, top, left and right. These can be set to the same value using C<set_border()> or individually using the relevant method calls shown above. Examples of the available border styles are shown in the 'Borders' worksheet created by formats.pl.




=head2 set_border_color()

    Also applies to:    set_bottom_color()
                        set_top_color()
                        set_left_color()
                        set_right_color()
                        
    Default state:      Color is off
    Default action:     Undefined
    Valid args:         See set_color()


Set the colour of the cell borders..




=head1 EXAMPLES

The following is a general example which demonstrates most of the features of the Spreadsheet::WriteExcel module:

    #!/usr/bin/perl -w
    
    use strict;
    use Spreadsheet::WriteExcel;
    
    # Create a new Excel workbook
    my $workbook = Spreadsheet::WriteExcel->new("regions.xls");
    
    # Add some worksheets
    my $north = $workbook->addworksheet("North");
    my $south = $workbook->addworksheet("South");
    my $east  = $workbook->addworksheet("East");
    my $west  = $workbook->addworksheet("West");
    
    # Add a Format
    my $format = $workbook->addformat();
    $format->set_bold();
    $format->set_color('blue');
    
    # Add a caption to each worksheet
    foreach my $worksheet (@{$workbook->worksheets()}) {
       $worksheet->write(0, 0, "Sales", $format);
    }
    
    # Write some data
    $north->write(0, 1, 200000);
    $south->write(0, 1, 100000);
    $east->write (0, 1, 150000);
    $west->write (0, 1, 100000);
    
    # Set the active worksheet
    $south->activate();
    
    # Set the width of the first column 
    $south->set_col_width(0, 0, 20);
    
    # Set the active cell
    $south->set_selection(0, 1);


The following example converts a tab separated file called C<tab.txt> into an Excel file called C<tab.xls>.

    #!/usr/bin/perl -w

    use strict;
    use Spreadsheet::WriteExcel;
    
    open (TABFILE, "tab.txt") or die "tab.txt: $!";
    
    my $workbook  = Spreadsheet::WriteExcel->new("tab.xls");
    my $worksheet = $workbook->addworksheet();
    
    # Row and column are zero indexed
    my $row = 0;
    
    while (<TABFILE>) {
        chomp;
        # Split on single tab
        my @Fld = split('\t', $_);
        
        my $col = 0;
        foreach my $token (@Fld) {
            $worksheet->write($row, $col, $token);
            $col++;
        }
        $row++;
    }

For an example of the range of fromatting options that are available with WriteExcel see the C<formats.pl> program in the main distribution.




=head1 LIMITATIONS

The following limits are imposed by Excel or the version of the BIFF file that has been implemented:

    Description                          Limit   Source
    -----------------------------------  ------  -------
    Maximum number of chars in a string  255     Excel 5
    Maximum number of columns            256     Excel 5, 97
    Maximum number of rows in Excel 5    16,384  Excel 5
    Maximum number of rows in Excel 97   65,536  Excel 97

The minimum file size is 6K due to the OLE overhead. The maximum file size is approximately 7MB (7087104 bytes) of BIFF data.




=head1 PORTABILITY

WriteExcel.pm will only work on systems where perl packs floats in 64 bit IEEE format. The float must also be in little-endian format but WriteExcel.pm will reverse it as necessary. Thus:

    print join(" ", map { sprintf "%#02x", $_ } unpack("C*", pack "d", 1.2345)), "\n";

should give (or in reverse order):

    0x8d 0x97 0x6e 0x12 0x83 0xc0 0xf3 0x3f


In general, if you don't know whether your system supports a 64 bit IEEE float or not, it probably does. If your system doesn't, WriteExcel will C<croak()> with the message given in the Diagnostics section.




=head1 DIAGNOSTICS

=over 4

=item Filename required in WriteExcel('Filename')

A filename must be given in the constructor.

=item Can't open filename. It may be in use.

The file cannot be opened for writing. It may be protected or already in use.

=item Required floating point format not supported on this platform.

Operating system doesn't support 64 bit IEEE float or it is byte-ordered in a way unknown to WriteExcel.


=item Maximum file size, 7087104, exceeded.

The current OLE implementation only supports a maximum BIFF file of this size.

=back




=head1 THE EXCEL BINARY FORMAT

Excel data is stored in the "Binary Interchange File Format" (BIFF) file format. Details of this format are given in the Excel SDK, the "Excel Developer's Kit" from Microsoft Press. It is also included in the MSDN CD library but is no longer available on the MSDN website. An older version of the BIFF documentation is available at http://www.cubic.org/source/archive/fileform/misc/excel.txt

Issues relating to the Excel SDK are discussed, occasionally, at news://microsoft.public.excel.sdk

The BIFF portion of the Excel file is comprised of contiguous binary records that have different functions and that hold different types of data. Each BIFF record is comprised of the following three parts:

        Record name;   Hex identifier, length = 2 bytes
        Record length; Length of following data, length = 2 bytes
        Record data;   Data, length = variable

The BIFF data is stored along with other data in an OLE Compound File. This is a structured storage which acts like a file system within a file. A Compound File is comprised of storages and streams which, to follow the file system analogy, are like directories and files.

The documentation for the OLE::Storage module, http://user.cs.tu-berlin.de/~schwartz/pmh/guide.html , contains one of the few descriptions of the OLE Compound File in the public domain.

For a open source implementation of the OLE library see the 'cole' library at http://arturo.directmail.org/filtersweb/

The source code for the Excel plugin of the Gnumeric spreadsheet also contains information relevant to the Excel BIFF format and the OLE container, http://www.gnumeric.org/

In addition the  source code for Star Office is available at http://www.openoffice.org/

Please note that the provision of this information does not constitute an invitation to start hacking at the BIFF or OLE file formats. There are more interesting ways to waste your time. ;)




=head1 WRITING EXCEL FILES

Depending on your requirements, background and general sensibilities you may prefer one of the following methods of getting data into Excel:

* CSV, comma separated variables or text. If the file extension is C<csv>, Excel will open and convert this format automatically.

* DBI, ADO or ODBC. Connect to an Excel file as a database. Using the appropriate driver Excel will behave like a database.

* HTML tables. This is an easy way of adding formatting.

* Win32::OLE module and office automation. See, the section "Reading Excel Files".




=head1 READING EXCEL FILES

Despite the title of this module the most commonly asked questions are in relation to reading Excel files. To read data from Excel files try:

* Spreadsheet::ParseExcel. This is a wrapper around the OLE::Storage module which makes it easy to extract data from an Excel file. http://search.cpan.org/search?dist=Spreadsheet-ParseExcel

* OLE::Storage, aka LAOLA. This is a Perl interface to OLE file formats. In particular, the distro contains an Excel to HTML converter called Herbert, http://user.cs.tu-berlin.de/~schwartz/pmh/

* DBI, ADO or ODBC. Connect to an Excel file as a database. Using the appropriate driver Excel will behave like a database.

* HTML tables. If the files are saved from Excel in a HTML format the data can be accessed using HTML::TableExtract http://search.cpan.org/search?dist=HTML-TableExtract

* Win32::OLE module and office automation. This requires a Windows platform and an installed copy of Excel. This is the most powerful and complete method for interfacing with Excel. See http://www.activestate.com/Products/ActivePerl/docs/faq/Windows/ActivePerl-Winfaq12.html and http://www.activestate.com/Products/ActivePerl/docs/site/lib/Win32/OLE.html

* There is also an open source C/C++ project based on the LAOLA work. Try the Filters Project at http://arturo.directmail.org/filtersweb/ and the xlHtml filter at the xlHtml Project at http://www.xlhtml.org/

If your main platform is UNIX but you have the resources to set up a separate Win32/MSOffice server, you can convert office documents to text, postscript or PDF using Win32::OLE. For a demonstration of how to do this using Perl see Docserver: http://search.cpan.org/search?mode=module&query=docserver

If you wish to view Excel files on a UNIX/Linux platform check out the excellent Gnumeric spreadsheet application at http://www.gnumeric.org/gnumeric or Star Office at http://www.openoffice.org/

If you wish to view Excel files on a Windows platform which doesn't have Excel installed you can use the free Microsoft Excel Viewer http://officeupdate.microsoft.com/2000/downloaddetails/xlviewer.htm


=head1 BUGS

There is a bug in the sequencing of the DESTROY methods of the contained objects in WriteExcel. This results in a file size of 0 bytes when the C<addworksheet()> or C<addformat()> methods are called in subroutines, ie. outside of the main scope of the program. This can be avoided by explicitly calling the C<close()> method for the WriteExcel object.

Orange isn't.

The Excel files that are produced by this module are not compatible with MS Access. Use DBI or ODBC instead.

The lack of a portable way of writing a little-endian 64 bit IEEE float.

QuickView: If you wish to write files that are fully compatible with QuickView it is necessary to write the cells in a sequential row by row order.




=head1 TO DO

Read a book, see a film or two.




=head1 ACKNOWLEDGEMENTS

The following people contributed to the debugging and testing of WriteExcel.pm:

Arthur@ais, Artur Silveira da Cunha, CPAN testers, Daniel Gardner, Harold Bamford, Johan Ekenberg, John Wren, Michael Buschauer, Mike Blazer, Paul J. Falbe.




=head1 AUTHOR

John McNamara jmcnamara@cpan.org

        I have known the inexorable sadness of pencils,
        Neat in their boxes, dolor of pad and paper weight,
        All the misery of manila folders and mucilage,
        Desolation in immaculate public places.
        Lonely reception room, lavatory, switch-board,
        The unalterable pathos of basin and pitcher,
        Ritual of multigraph, paper clip, comma,
        Endless duplication of lives and objects.
        And I have seen dust from the walls of institutions,
        Finer than flour, alive, more dangerous than silica,
        Sift, almost invisible, through long afternoons of tedium,
        Dropping a fine film on nails and delicate eyebrows,
        Glazing the pale hair, the duplicate grey standard faces.
        
        - Theodore Roethke




=head1 COPYRIGHT

Copyright (c) 2000, John McNamara. All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

