package Spreadsheet::WriteExcel;

###############################################################################
#
# WriteExcel.
#
# Spreadsheet::WriteExcel - Write to a cross-platform Excel binary file.
#
# Copyright 2000-2001, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

require Exporter;

use strict;
use Spreadsheet::WriteExcel::Workbook;



use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::WriteExcel::Workbook Exporter);

$VERSION = '0.30'; # 26 February 2001, O'Hara

###############################################################################
#
# new()
#
# Constructor. Wrapper for a Workbook object.
# uses: Spreadsheet::WriteExcel::BIFFwriter
#       Spreadsheet::WriteExcel::OLEwriter
#       Spreadsheet::WriteExcel::Workbook
#       Spreadsheet::WriteExcel::Worksheet
#       Spreadsheet::WriteExcel::Format
#       Spreadsheet::WriteExcel::Formula
#
sub new {

    my $class = shift;
    my $self  = Spreadsheet::WriteExcel::Workbook->new(@_);

    bless  $self, $class;
    return $self;
}


1;


__END__



=head1 NAME

Spreadsheet::WriteExcel - Write to a cross-platform Excel binary file.




=head1 VERSION

This document refers to version 0.30 of Spreadsheet::WriteExcel, released February 26, 2001.




=head1 SYNOPSIS

To write a string, a formatted string, a number and a formula to the first worksheet in an Excel workbook called perl.xls:

    use Spreadsheet::WriteExcel;

    # Create a new Excel workbook
    my $workbook = Spreadsheet::WriteExcel->new("perl.xls");
    $worksheet   = $workbook->addworksheet(); # Add a worksheet
    $format      = $workbook->addformat();    # Add a format
    
    # Define the format
    $format->set_bold();
    $format->set_color('red');
    $format->set_align('center');
    
    # Row and column are zero indexed
    $col = 0;
    $row = 0;

    $worksheet->write($row,  $col, "Hi Excel!");
    $worksheet->write(1,     $col, "Hi Excel!", $format);
    $worksheet->write(2,     $col, 1.2345);
    $worksheet->write(3,     $col, '=SIN(PI()/4)');




=head1 DESCRIPTION

The Spreadsheet::WriteExcel module can be used create a cross-platform Excel binary file. Multiple worksheets can be added to a workbook and formatting can be applied to cells. Text, numbers, formulas and hyperlinks can be written to the cells. 

The Excel file produced by this module is compatible with Excel 5, 95, 97 and 2000.

The module will work on the majority of Windows, UNIX and Macintosh platforms. Generated files are also compatible with the Linux/UNIX spreadsheet applications OpenOffice, Gnumeric and XESS. The generated files are not compatible with MS Access.




=head1 WORKBOOK METHODS

The Spreadsheet::WriteExcel module provides an object oriented interface to a new Excel workbook. The following methods are available through a new workbook.

If you are unfamiliar with object oriented interfaces or the way that they are implemented in Perl have a look at C<perlobj> and C<perltoot> in the main Perl documentation.


=head2 new()

A new Excel workbook is created using the C<new()> constructor as follows:

    my $workbook = Spreadsheet::WriteExcel->new("filename.xls");

Although C<my> is not specifically required it defines the scope of the new workbook variable and, in the majority of cases, ensures that the workbook is closed properly without explicitly calling the C<close()> method.

You can redirect the output to STDOUT using the special Perl filehandle C<"-">. This can be useful for CGIs which have a Content-type of C<application/vnd.ms-excel>, for example:

    #!/usr/bin/perl -w

    use strict;
    use Spreadsheet::WriteExcel;

    # Set the filename and send the content type and disposition
    my $filename ="cgitest.xls";

    print "Content-type: application/vnd.ms-excel\n";
    print "Content-Disposition: attachment; filename=$filename\n";
    print "\n";


    my $workbook  = Spreadsheet::WriteExcel->new("-");
    my $worksheet = $workbook->addworksheet();
    
    $workbook->write(0, 0, "Hi Excel!");




=head2 close()

The C<close()> method can be called to explicitly close an Excel file.

    $workbook->close();

An explicit C<close()> is required if the file must be closed prior to performing some external action on it such as copying or reading its size.

In addition, C<close()> may be required if the scope of the Workbook, Worksheet or Format objects cannot be determined by perl. Situations where this can occur are:

=over

=item * If C<my()> was not used to declare the scope of a workbook variable created using C<new()>.

=item * If the C<addworksheet()> or C<addformat()> methods are called in subroutines.

=back

The reason for this is that Spreadsheet::WriteExcel relies on Perl's C<DESTROY> subroutine to trigger destructor methods in a specific sequence. This will not happen if the scope of the variables cannot be determined.


In general, if you create a file with a size of 0 bytes you need to call C<close()>.




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




=head2 worksheets()

The C<worksheets()> method returns a reference to the array of worksheets in a workbook. This can be useful if you want to repeat an operation on each worksheet in a workbook or if you wish to refer to a worksheet by its index:

    foreach $worksheet (@{$workbook->worksheets()}) {
       $worksheet->write(0, 0, "Hello");
    }
    
    # or:
    
    $worksheets = $workbook->worksheets();
    @$worksheets[0]->write(0, 0, "Hello");


References are explained in detail in C<perlref> and C<perlreftut> in the main Perl documentation.




=head2 set_1904()

Excel stores dates as real numbers where the integer part stores the number of days since the epoch and the fractional part stores the percentage of the day. The epoch can be either 1900 or 1904. Excel for Windows uses 1900 and Excel for Macintosh uses 1904. However, Excel on either platform will convert automatically between one system and the other.

Spreadsheet::WriteExcel stores dates in the 1900 format by default. If you wish to change this you can call the C<set_1904()> workbook method. You can query the current value by calling the C<get_1904()> workbook method. This returns 0 for 1900 and 1 for 1904.

In general you probably won't need to use C<set_1904()>.




=head1 WORKSHEET METHODS

The following methods are available through a new worksheet. A new worksheet is created by calling the C<addworksheet()> method from a workbook object:

    $worksheet1 = $workbook->addworksheet();
    $worksheet2 = $workbook->addworksheet();




=head2 write($row, $column, $token, $format)

The C<write()> method is a general alias for one of several methods of writing to a cell in Excel. C<write()> calls one of the following methods depending on the value of C<$token>:

C<write_number()> if C<$token> is a number based on the following regex: C<$token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/>.

C<write_blank()> if C<$token> is a blank string: C<""> or C<''>.

C<write_url()> if C<$token> is a URL based on the following regex: C<$token =~ m|[fh]tt?p://|>.

C<write_formula()> if the first character of C<$token> is C<=>.

C<write_string()> if none of the previous conditions apply.

Here are some examples:

    $worksheet->write(0, 0, "Hello"                 );  # write_string()
    $worksheet->write(1, 0, 'One'                   );  # write_string()
    $worksheet->write(2, 0,  2                      );  # write_number()
    $worksheet->write(3, 0,  3.00001                );  # write_number()
    $worksheet->write(4, 0,  ""                     );  # write_blank()
    $worksheet->write(5, 0,  ''                     );  # write_blank()
    $worksheet->write(6, 0,  'http://www.perl.com/' );  # write_url()
    $worksheet->write(7, 0,  'ftp://ftp.cpan.org/'  );  # write_url()
    $worksheet->write(8, 0,  '=A3 + 3*A4'           );  # write_formula()
    $worksheet->write(9, 0,  '=SIN(PI()/4)'         );  # write_formula()

The C<$format> argument is optional. It should be a valid Format object, see L<FORMAT METHODS>:

    my $format = $workbook->addformat();
    $format->set_bold();
    $format->set_color('red');
    $format->set_align('center');

    $worksheet->write(4, 0, "Hello", $format ); # Formatted string


It should be noted that C<$row> and C<$column> are zero indexed cell locations for the C<write> methods. Thus, Cell A1 is (0, 0) and Cell AD2000 is (1999, 29). Cells can be written to in any order but for forward compatibility it is probably best to write them in row-column order when possible.

The C<write> methods return:

    0 for success.
   -1 for insufficient number of arguments.
   -2 for row or column out of bounds.
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




=head2 write_formula($row, $column, $formula, $format)

Write a formula or function to the cell specified by C<$row> and C<$column>:

    $worksheet->write_formula(0, 0, '=$B$3 + B4'  );
    $worksheet->write_formula(1, 0, '=SIN(PI()/4)');
    $worksheet->write_formula(2, 0, '=SUM(B1:B5)');
    $worksheet->write_formula(3, 0, '=IF(A3>1,"Yes", "No")');
    $worksheet->write_formula(4, 0, '=AVERAGE(1, 2, 3, 4)');
    $worksheet->write_formula(5, 0, '=DATEVALUE("1-Jan-2001")');

For more information about writing Excel formulas see L<FORMULAS AND FUNCTIONS IN EXCEL>




=head2 write_blank($row, $column, $format)

Write a blank cell specified by C<$row> and C<$column>:

    $worksheet->write_blank(0, 0, $format);

This method is useful for adding formatting to a cell which doesn't contain a string or number value.




=head2 write_url($row, $col, $url, $string, $format)

Write a hyperlink to a URL in the cell specified by C<$row> and C<$column>. The hyperlink is comprised of two elements: the visible label and the invisible link. The visible label is the same as the link unless an alternative string is specified. The alternative C<$string> and the C<$format> are optional. 

    $worksheet->write_url(0, 0, 'http://www.perl.com/'                 );
    $worksheet->write_url(1, 0, 'http://www.perl.com/', 'Perl home'    );
    $worksheet->write_url(2, 0, 'http://www.perl.com/', undef, $format );

The label is written using the C<write_string()> method. Therefore the 255 characters string limit applies to the label: the URL can be any length. Use C<undef> if you wish to specify a format without specifying an alternative string.

Note: Hyperlinks are not available in Excel 5. They will appear as a string only.


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

    $worksheet1->set_selection(3, 3);       # Cell D4
    $worksheet2->set_selection(3, 3, 6, 6); # Cells D4 to G7
    $worksheet3->set_selection(6, 6, 3, 3); # Cells G7 to D4

The default is cell (0, 0).




=head2 set_row($row, $height, $format);

This method can be used to specify the height of a row. The C<$format> argument is optional, for additional information, see L<FORMAT METHODS>.

    $worksheet->set_row(0, 20); # Row 1 height set to 20

If you wish to set the format without changing the height you can pass C<undef> as the height parameter:

    $worksheet->set_row(0, undef, $height);


=head2 set_column($first_col, $last_col, $width, $format);

This method can be used to specify the width of a single column or a range of columns. If the method is applied to a single column the value of C<$first_col> and C<$last_col> should be the same:

    $worksheet->set_column(0, 0, 20); # Column A width set to 20
    $worksheet->set_column(1, 3, 30); # Columns B-D width set to 30

The width corresponds to the column width value that is specified in Excel. It is approximately equal to the length of a string in the default font of Arial 10. The C<$format> argument is optional, for additional information, see L<FORMAT METHODS>.

If you wish to set the format without changing the width you can pass C<undef> as the width parameter:

    $worksheet->set_column(0, 0, undef, $format);




=head1 FORMAT METHODS

This section describes the methods that are available through a Format object. Format objects are created by calling the workbook C<addformat()> method as follows:

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
    
    Number          Numeric format      set_num_format()
    
    Alignment       Horizontal align    set_align()
                    Vertical align      set_align()
                    Rotation            set_rotation()
                    Text wrap           set_text_wrap()
                    Justify last        set_text_justlast()
                    Merge               set_merge()
    
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

Once a Format object has been constructed it can be passed as an argument to the worksheet C<write> methods as follows:

    $worksheet->write(0, 0, "One", $format);
    $worksheet->write_string(1, 0, "Two", $format);
    $worksheet->write_number(2, 0, 3, $format);
    $worksheet->write_blank(3, 0, $format);

Formats can also be passed to the worksheet C<set_row()> and C<set_column()> methods to define the default property for a row or column.

    $worksheet->set_row(0, 15, $format);
    $worksheet->set_column(0, 0, 15, $format);

However, the C<set_row()> and C<set_column()> methods will not set the format for individual cells written by WriteExcel, they only have an effect on cells written after the workbook is opened in Excel.


It is important to understand that a Format is applied to a cell not in its current state but in its final state. Consider the following example:

    my $format = $workbook->addformat();
    $format->set_bold();
    $format->set_color('red');
    $worksheet->write(0, 0, "Cell A1", $format);
    $format->set_color('green');
    $worksheet->write(0, 1, "Cell B1", $format);

Cell A1 is assigned the Format C<$format> which is initially set to the colour red. However, the colour is subsequently set to green. When Excel displays Cell A1 it will display the final state of the Format which in this case will be the colour green.

The Format object methods are described in more detail in the following sections. In addition, there is a Perl program called C<formats.pl> in the C<examples> directory of the WriteExcel distribution. This program creates an Excel workbook called C<formats.xls> which contains examples of almost all the format types.




=head2 copy($format)


This is the only method of a Format object that doesn't apply directly to a property. It is used to copy all of the properties from one Format object to another:

    my $lorry1 = $workbook->addformat();
    $lorry1->set_bold();
    $lorry1->set_italic();
    $lorry1->set_color('red');    # lorry1 is bold, italic and red

    my $lorry2 = $workbook->addformat();
    $lorry2->copy($lorry1);
    $lorry2->set_color('yellow'); # lorry2 is bold, italic and yellow

This can be useful when you are setting up several complex but similar formats. It is also useful if you want to use a format in more than one workbook:

    # Create the workbooks
    my $workbook1   = Spreadsheet::WriteExcel->new("workbook1.xls");
    my $workbook2   = Spreadsheet::WriteExcel->new("workbook2.xls");
    my $worksheet1  = $workbook->addworksheet();
    my $worksheet2  = $workbook->addworksheet();
    my $format1     = $workbook->addformat();
    my $format2     = $workbook->addformat();
    
    # Create a global format object that isn't tied to a workbook
    my $global      = Spreadsheet::WriteExcel::Format->new();
    $global->set_color('blue');
    
    # Copy the global format properties to the worksheet formats
    $format1->copy($global);
    $format2->copy($global);

Note: this is not a copy constructor, both objects must exist prior to copying.




=head2 set_font($fontname)

    Default state:      Font is Arial
    Default action:     None
    Valid args:         Any valid font name

Specify the font used:

    $format->set_font('Times New Roman');

Excel can only display fonts that are installed on the system that it is running on. Therefore it is best to use the fonts that come as standard such as 'Arial', 'Times New Roman' and 'Courier New'. See also the Fonts worksheet created by formats.pl




=head2 set_size()

    Default state:      Font size is 10
    Default action:     Set font size to 1
    Valid args:         Integer values from 1 to as big as your screen.


Set the font size. Excel adjusts the height of a row to accommodate the largest font size in the row. You can also explicitly specify the height of a row using the set_row() worksheet method.

    my $format = $workbook->addformat();
    $format->set_size(30);





=head2 set_color()

    Default state:      Excels default color, usually black
    Default action:     Set the default color
    Valid args:         Integers form 8..63 or the following strings:
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

Set the font colour. The C<set_color()> method is used as follows:

    my $format = $workbook->addformat();
    $format->set_color('red');
    $worksheet->write(0, 0, "wheelbarrow", $format);

Note: The C<set_color()> method is used to set the colour of the font in a cell. To set the colour of a cell use the C<set_fg_color()> and C<set_pattern()> methods.

For additional examples see the 'Named colors' and 'Standard colors' worksheets created by formats.pl




=head2 set_bold()

    Default state:      bold is off
    Default action:     Turn bold on
    Valid args:         0, 1 [1]

Set the bold property of the font: 

    $format->set_bold();  # Turn bold on

[1] Actually, values in the range 100..1000 are also valid. 400 is normal, 700 is bold and 1000 is very bold indeed. It is probably best to set the value to 1 and use normal bold.




=head2 set_italic()

    Default state:      Italic is off
    Default action:     Turn italic on
    Valid args:         0, 1

Set the italic property of the font.




=head2 set_underline()

    Default state:      Underline is off
    Default action:     Turn on single underline
    Valid args:         0  = No underline
                        1  = Single underline
                        2  = Double underline
                        33 = Single accounting underline
                        34 = Double accounting underline

Set the underline property of the font.




=head2 set_strikeout()

    Default state:      Strikeout is off
    Default action:     Turn strikeout on
    Valid args:         0, 1

Set the strikeout property of the font.




=head2 set_script()

    Default state:      Super/Subscript is off
    Default action:     Turn Superscript on
    Valid args:         0  = Normal
                        1  = Superscript
                        2  = Subscript

Set the superscript/subscript property of the font. This format is currently not very useful.




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




=head2 set_num_format()

    Default state:      General format
    Default action:     Format index 1
    Valid args:         See the following table

This method is used to define the numerical format of a number in Excel. It controls whether a number is displayed as an integer, a floating point number, a date, a currency value or some other user defined format.

The numerical format of a cell can be specified by using a format string or an index to one of Excel's built-in formats:

    my $format1 = $workbook->addformat();
    my $format2 = $workbook->addformat();
    $format1->set_num_format('d mmm yyyy'); # Format string
    $format2->set_num_format(0x0f);         # Format index
    
    $worksheet->write(0, 0, 36892.521, $format1);      # 1 Jan 2001
    $worksheet->write(0, 0, 36892.521, $format2);      # 1-Jan-01


Using format strings you can define very sophisticated formatting of numbers.


    $format01->set_num_format('0.000');
    $worksheet->write(0,  0, 3.1415926, $format01);    # 3.142

    $format02->set_num_format('#,##0');
    $worksheet->write(1,  0, 1234.56,   $format02);    # 1,235

    $format03->set_num_format('#,##0.00');
    $worksheet->write(2,  0, 1234.56,   $format03);    # 1,234.56

    $format04->set_num_format('$0.00');
    $worksheet->write(3,  0, 49.99,     $format04);    # $49.99

    $format05->set_num_format('£0.00');
    $worksheet->write(4,  0, 49.99,     $format05);    # £49.99

    $format06->set_num_format('¥0.00');
    $worksheet->write(5,  0, 49.99,     $format06);    # ¥49.99

    $format07->set_num_format('mm/dd/yy');
    $worksheet->write(6,  0, 36892.521, $format07);    # 01/01/01

    $format08->set_num_format('mmm d yyyy');
    $worksheet->write(7,  0, 36892.521, $format08);    # Jan 1 2001

    $format09->set_num_format('d mmmm yyyy');
    $worksheet->write(8,  0, 36892.521, $format09);    # 1 January 2001

    $format10->set_num_format('dd/mm/yyyy hh:mm AM/PM');
    $worksheet->write(9,  0, 36892.521, $format10);    # 01/01/2001 12:30 AM

    $format11->set_num_format('0 "dollar and" .00 "cents"');
    $worksheet->write(10, 0, 1.87,      $format11);    # 1 dollar and .87 cents

    # Conditional formatting
    $format12->set_num_format('[Green]General;[Red]-General;General');
    $worksheet->write(11, 0, 123,       $format12);    # > 0 Green
    $worksheet->write(12, 0, -45,       $format12);    # < 0 Red
    $worksheet->write(13, 0, 0,         $format12);    # = 0 Default colour

The colour format should have one of the following values:

    [Black] [Blue] [Cyan] [Green] [Magenta] [Red] [White] [Yellow]

Alternatively you can specify the colour based on a colour index as follows: C<[Color n]>, where n is a standard Excel colour index - 7. See the 'Standard colors' worksheet created by formats.pl.

For more information refer to the documentation on formatting in the C<doc> directory of the Spreadsheet::WriteExcel distro, the Excel on-line help or to the tutorial at: http://support.microsoft.com/support/Excel/Content/Formats/default.asp and http://support.microsoft.com/support/Excel/Content/Formats/codes.asp

You should ensure that the format string is valid in Excel prior to using it in WriteExcel.

One of the most common uses of the C<set_num_format()> is to format a number as a date. Excel stores dates as a real number where the integer part of the number stores the number of days since the epoch and the fractional part stores the percentage of the day. The epoch can be either 1900 or 1904. Excel for Windows uses 1900 and Excel for Macintosh uses 1904. However, Excel on either platform will convert automatically between one system and the other. For an example of how to convert between UNIX/Perl time and Excel time have a look at the C<ms_time.pl> program in the C<examples> directory of the WriteExcel distribution.


Excel's built-in formats are shown in the following table:

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


For examples of these formatting codes see the 'Numerical formats' worksheet created by formats.pl.


Note 1. Numeric formats 23 to 36 are not documented by Microsoft and may differ in international versions.

Note 2. In Excel 5 the dollar sign appears as a dollar sign. In Excel 97-2000 it appears as the defined local currency symbol.

Note 3. The red negative numeric formats display slightly differently in Excel 5 and Excel 97-2000.




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
    $worksheet->set_row(0, 30);
    $worksheet->write(0, 0, "X", $format);

Text can be aligned across two or more adjacent cells using the C<merge> property. See also, the C<set_merge()> method.

The C<vjustify> (vertical justify) option can be used to provide automatic text wrapping in a cell. The height of the cell will be adjusted to accommodate the wrapped text. To specify where the text wraps use the C<set_text_wrap()> method.


For further examples see the 'Alignment' worksheet created by formats.pl.




=head2 set_merge()

    Default state:      Cell merging is off
    Default action:     Turn cell merging on
    Valid args:         1

Text can be aligned across two or more adjacent cells using the C<set_merge()> method. This is an alias for the unintuitive C<set_align('merge')> method call.

Only one cell should contain the text, the other cells should be blank:

    my $format = $workbook->addformat();
    $format->set_merge();

    $worksheet->write(1, 1, 'Merged cells', $format);
    $worksheet->write_blank(1, 2, $format);

See also the C<merge1.pl> and C<merge2.pl> programs in the C<examples> directory.



=head2 set_text_wrap()

    Default state:      Text wrap is off
    Default action:     Turn text wrap on
    Valid args:         0, 1


Here is an example using the text wrap property, the escape character C<\n> is used to indicate the end of line:

    my $format = $workbook->addformat();
    $format->set_text_wrap();
    $worksheet->write(0, 0, "It's\na bum\nwrap", $format);

Excel will adjust the height of the row to accommodate the wrapped text. A similar effect can be obtained without newlines using the C<set_align('vjustify')> method. See the C<textwrap.pl> program in the C<examples> directory.



=head2 set_rotation()

    Default state:      Text rotation is off
    Default action:     Rotation style 1
    Valid args:         0 No rotation
                        1 Letters run from top to bottom
                        2 90° anticlockwise
                        3 90° clockwise


Set the rotation of the text in a cell. See the 'Alignment' worksheet created by formats.pl.





=head2 set_text_justlast()

    Default state:      Justify last is off
    Default action:     Turn justify last on
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


Note, the foreground and background colours will only have an effect if the cell pattern has been set. In the most common case you can specify the solid fill pattern and the foreground colour as follows:

    my $format = $workbook->addformat();
    $format->set_pattern();         # Set pattern to 1, i.e. solid fill
    $format->set_fg_color('green'); # Note foreground and not background
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


A cell border is comprised of a border on the bottom, top, left and right. These can be set to the same value using C<set_border()> or individually using the relevant method calls shown above. Examples of the available border styles are shown in the 'Borders' worksheet created by formats.pl.




=head2 set_border_color()

    Also applies to:    set_bottom_color()
                        set_top_color()
                        set_left_color()
                        set_right_color()
                        
    Default state:      Color is off
    Default action:     Undefined
    Valid args:         See set_color()


Set the colour of the cell borders.




=head1 FORMULAS AND FUNCTIONS IN EXCEL

The first thing to note is that there are still some outstanding issues with the implementation of formulas and functions:

    * Writing a formula is much slower than writing the equivalent string. 
    * Unary minus isn't supported.
    * You cannot reference cells outside of the current worksheet.
    * You cannot use arrays constants, i.e. {1;2;3}, in functions.
    * You cannot use embedded double quotes in strings.
    * Whitespace is not preserved around operators.

However, these constraints will be removed in future versions. They are here because of a trade-off between features and time.

The following is a brief introduction to formulas and functions in Excel and Spreadsheet::WriteExcel.

A formula is a string that begins with an equal sign:

    '=A1+B1'
    '=AVERAGE(1, 2, 3)'

The formula can contain numbers, strings, boolean values, cell references, cell ranges and functions. Formulas should be written as they appear in Excel, that is cells and functions must be in uppercase.

Cells in Excel are referenced using the A1 notation system where the column is designated by a letter and the row by a number. Columns range from A to IV i.e. 0 to 255, rows range from 1 to 16384. This system is different from the zero indexed row and column syntax used by the Spreadsheet::WriteExcel C<write()> methods where, for example, the cell C<C7> would be referred to as C<(6,2)>. In the C<examples> directory of the distro there is program called C<A1convert.pl> which contains functions to help you convert between the two systems.

The Excel C<$> notation in cell references is also supported. This allows you to specify whether a row or column is relative or absolute. This only has an effect if the cell is copied. The following examples show relative and absolute values.

    '=A1'   # Column and row are relative
    '=$A1'  # Column is absolute and row is relative
    '=A$1'  # Column is relative and row is absolute
    '=$A$1' # Column and row are absolute

The following table lists the operators that are available in Excel functions. The majority of the operators are the same as Perl's. Differences are indicated.

    Arithmetic operators:
    =====================
    Operator    Meaning                     Example 
    
       +        Addition                    1+2 
       -        Subtraction                 2-1
       *        Multiplication              2*3 
       /        Division                    1/4 
       ^        Exponentiation              2^3     # Equivalent to **
       -        Unary minus                 -(1+2)  # Not yet supported
       %        Percent (Not modulus)       13%     # Not supported, Note [1]
    
    
    Comparison operators:
    =====================
    Operator    Meaning                     Example  
        =       Equal to                    A1 =  B1    # Equivalent to ==
        <>      Not equal to                A1 <> B1    # Equivalent to !=
        >       Greater than                A1 >  B1 
        <       Less than                   A1 <  B1 
        >=      Greater than or equal to    A1 > =B1 
        <=      Less than or equal to       A1 <= B1 
    
    
    String operator:
    ================
    Operator    Meaning                     Example  
    
        &       Concatenation               "Hello " & "World!"  # Note [2]
    
    
    Reference operators:
    ====================
    Operator    Meaning                     Example  
        :       Range operator              A1:A4               # Note [3]
        ,       Union operator              SUM(1, 2+2, B3)     # Note [4]   
    
    
    Note [1]: You can get a percentage with formatting and modulus with MOD().
    Note [2]: Equivalent to ("Hello " . "World!") in Perl.
    Note [3]: This range is equivalent to cells A1, A2, A3 and A4.
    Note [4]: The comma behaves like the list separator in Perl.

The range and comma operators can have different symbols in non-English versions of Excel. These will be supported in a later version of Spreadsheet::WriteExcel.

The following table lists all of the core functions supported by Excel 5 and Spreadsheet::WriteExcel. Any additional functions that are available through the "Analysis ToolPak" or other add-ins are not supported. 

    ABS           DB            INDIRECT      NORMINV       SLN
    ACOS          DCOUNT        INFO          NORMSDIST     SLOPE
    ACOSH         DCOUNTA       INT           NORMSINV      SMALL
    ADDRESS       DDB           INTERCEPT     NOT           SQRT
    AND           DEGREES       IPMT          NOW           STANDARDIZE
    AREAS         DEVSQ         IRR           NPER          STDEV
    ASIN          DGET          ISBLANK       NPV           STDEVP
    ASINH         DMAX          ISERR         ODD           STEYX
    ATAN          DMIN          ISERROR       OFFSET        SUBSTITUTE
    ATAN2         DOLLAR        ISLOGICAL     OR            SUBTOTAL
    ATANH         DPRODUCT      ISNA          PEARSON       SUM
    AVEDEV        DSTDEV        ISNONTEXT     PERCENTILE    SUMIF
    AVERAGE       DSTDEVP       ISNUMBER      PERCENTRANK   SUMPRODUCT
    BETADIST      DSUM          ISREF         PERMUT        SUMSQ
    BETAINV       DVAR          ISTEXT        PI            SUMX2MY2
    BINOMDIST     DVARP         KURT          PMT           SUMX2PY2
    CALL          ERROR.TYPE    LARGE         POISSON       SUMXMY2
    CEILING       EVEN          LEFT          POWER         SYD
    CELL          EXACT         LEN           PPMT          T
    CHAR          EXP           LINEST        PROB          TAN
    CHIDIST       EXPONDIST     LN            PRODUCT       TANH
    CHIINV        FACT          LOG           PROPER        TDIST
    CHITEST       FALSE         LOG10         PV            TEXT
    CHOOSE        FDIST         LOGEST        QUARTILE      TIME
    CLEAN         FIND          LOGINV        RADIANS       TIMEVALUE
    CODE          FINV          LOGNORMDIST   RAND          TINV
    COLUMN        FISHER        LOOKUP        RANK          TODAY
    COLUMNS       FISHERINV     LOWER         RATE          TRANSPOSE
    COMBIN        FIXED         MATCH         REGISTER.ID   TREND
    CONCATENATE   FLOOR         MAX           REPLACE       TRIM
    CONFIDENCE    FORECAST      MDETERM       REPT          TRIMMEAN
    CORREL        FREQUENCY     MEDIAN        RIGHT         TRUE
    COS           FTEST         MID           ROMAN         TRUNC
    COSH          FV            MIN           ROUND         TTEST
    COUNT         GAMMADIST     MINUTE        ROUNDDOWN     TYPE
    COUNTA        GAMMAINV      MINVERSE      ROUNDUP       UPPER
    COUNTBLANK    GAMMALN       MIRR          ROW           VALUE
    COUNTIF       GEOMEAN       MMULT         ROWS          VAR
    COVAR         GROWTH        MOD           RSQ           VARP
    CRITBINOM     HARMEAN       MODE          SEARCH        VDB
    DATE          HLOOKUP       MONTH         SECOND        VLOOKUP
    DATEVALUE     HOUR          N             SIGN          WEEKDAY
    DAVERAGE      HYPGEOMDIST   NA            SIN           WEIBULL
    DAY           IF            NEGBINOMDIST  SINH          YEAR
    DAYS360       INDEX         NORMDIST      SKEW          ZTEST

Function names in other European languages will be supported in a later version of Spreadsheet::WriteExcel. For a general introduction to Excel's formulas and an explanation of the functions have a look at the Excel help files or the following links: http://msdn.microsoft.com/library/default.asp?URL=/library/officedev/office97/s88f2.htm and http://msdn.microsoft.com/library/default.asp?URL=/library/officedev/office97/s992f.htm


If your formula doesn't work in Spreadsheet::WriteExcel try the following:

    1. Verify that the formula works in Excel (or Gnumeric or OpenOffice).
    2. Ensure that it isn't on the TODO list at the start of this section.
    3. Ensure that cell references and cell names are in uppercase.
    4. Ensure that you are using the U.S. style range and union operators.
    5. Ensure the function is in the above table.

If you go through steps 1-5 and you still have a problem, mail me.




=head1 EXAMPLES


There are additional examples in the C<examples> directory of the Spreadsheet::WriteExcel distro.




=head2 Example 1 

The following example shows some of the basic features of Spreadsheet::WriteExcel.


    #!/usr/bin/perl -w
    
    use strict;
    use Spreadsheet::WriteExcel;

    # Create a new workbook called simple.xls and add a worksheet
    my $workbook  = Spreadsheet::WriteExcel->new("simple.xls");
    my $worksheet = $workbook->addworksheet();

    # The general syntax is write($row, $column, $token). Note that row and
    # column are zero indexed

    # Write some text
    $worksheet->write(0, 0,  "Hi Excel!");


    # Write some numbers
    $worksheet->write(2, 0,  3);          # Writes 3
    $worksheet->write(3, 0,  3.00000);    # Writes 3
    $worksheet->write(4, 0,  3.00001);    # Writes 3.00001
    $worksheet->write(5, 0,  3.14159);    # TeX revision no.?


    # Write some formulas
    $worksheet->write(7, 0,  '=A3 + A6');
    $worksheet->write(8, 0,  '=IF(A5>3,"Yes", "No")');


    # Write a hyperlink
    $worksheet->write(10, 0, 'http://www.perl.com/');




=head2 Example 2 

The following is a general example which demonstrates some features of working with multiple worksheets.

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
    $south->set_column(0, 0, 20);
    
    # Set the active cell
    $south->set_selection(0, 1);




=head2 Example 3 

This example shows how to use a conditional numerical format with colours to indicate if a share price has gone up or down.

    use strict;
    use Spreadsheet::WriteExcel;

    # Create a new workbook and add a worksheet
    my $workbook  = Spreadsheet::WriteExcel->new("stocks.xls");
    my $worksheet = $workbook->addworksheet();

    # Set the column width for columns 1, 2, 3 and 4
    $worksheet->set_column(0, 3, 15);


    # Create a format for the column headings
    my $header = $workbook->addformat();
    $header->set_bold();
    $header->set_size(12);
    $header->set_color('blue');


    # Create a format for the stock price
    my $f_price = $workbook->addformat();
    $f_price->set_align('left');
    $f_price->set_num_format('$0.00');


    # Create a format for the stock volume
    my $f_volume = $workbook->addformat();
    $f_volume->set_align('left');
    $f_volume->set_num_format('#,##0');


    # Create a format for the price change. This is an example of a conditional
    # format. The number is formatted as a percentage. If it is positive it is
    # formatted in green, if it is negative it is formatted in red and if it is
    # zero it is formatted as the default font colour (in this case black).
    # Note: the [Green] format produces an unappealing lime green. Try 
    # [Color 10] instead for a dark green.
    #
    my $f_change = $workbook->addformat();
    $f_change->set_align('left');
    $f_change->set_num_format('[Green]0.0%;[Red]-0.0%;0.0%');


    # Write out the data
    $worksheet->write(0, 0, 'Company',$header);
    $worksheet->write(0, 1, 'Price',  $header);
    $worksheet->write(0, 2, 'Volume', $header);
    $worksheet->write(0, 3, 'Change', $header);

    $worksheet->write(1, 0, 'Damage Inc.'       );
    $worksheet->write(1, 1, 30.25,    $f_price ); # $30.25
    $worksheet->write(1, 2, 1234567,  $f_volume); # 1,234,567
    $worksheet->write(1, 3, 0.085,    $f_change); # 8.5% in green

    $worksheet->write(2, 0, 'Dump Corp.'        );
    $worksheet->write(2, 1, 1.56,     $f_price ); # $1.56
    $worksheet->write(2, 2, 7564,     $f_volume); # 7,564
    $worksheet->write(2, 3, -0.015,   $f_change); # -1.5% in red

    $worksheet->write(3, 0, 'Rev Ltd.'          );
    $worksheet->write(3, 1, 0.13,     $f_price ); # $0.13
    $worksheet->write(3, 2, 321,      $f_volume); # 321
    $worksheet->write(3, 3, 0,        $f_change); # 0 in the font color (black)




=head2 Example 4 

The following is a simple example of using functions.

    #!/usr/bin/perl -w
    
    use strict;
    use Spreadsheet::WriteExcel;

    # Create a new workbook and add a worksheet
    my $workbook  = Spreadsheet::WriteExcel->new("stats.xls");
    my $worksheet = $workbook->addworksheet('Test data');

    # Set the column width for columns 1
    $worksheet->set_column(0, 0, 20);


    # Create a format for the headings
    my $format = $workbook->addformat();
    $format->set_bold();


    # Write the sample data
    $worksheet->write(0, 0, 'Sample', $format);
    $worksheet->write(0, 1, 1);
    $worksheet->write(0, 2, 2);
    $worksheet->write(0, 3, 3);
    $worksheet->write(0, 4, 4);
    $worksheet->write(0, 5, 5);
    $worksheet->write(0, 6, 6);
    $worksheet->write(0, 7, 7);
    $worksheet->write(0, 8, 8);

    $worksheet->write(1, 0, 'Length', $format);
    $worksheet->write(1, 1, 25.4);
    $worksheet->write(1, 2, 25.4);
    $worksheet->write(1, 3, 24.8);
    $worksheet->write(1, 4, 25.0);
    $worksheet->write(1, 5, 25.3);
    $worksheet->write(1, 6, 24.9);
    $worksheet->write(1, 7, 25.2);
    $worksheet->write(1, 8, 24.8);

    # Write some statistical functions
    $worksheet->write(4,  0, 'Count', $format);
    $worksheet->write(4,  1, '=COUNT(B1:I1)');

    $worksheet->write(5,  0, 'Sum', $format);
    $worksheet->write(5,  1, '=SUM(B2:I2)');

    $worksheet->write(6,  0, 'Average', $format);
    $worksheet->write(6,  1, '=AVERAGE(B2:I2)');

    $worksheet->write(7,  0, 'Min', $format);
    $worksheet->write(7,  1, '=MIN(B2:I2)');

    $worksheet->write(8,  0, 'Max', $format);
    $worksheet->write(8,  1, '=MAX(B2:I2)');

    $worksheet->write(9,  0, 'Standard Deviation', $format);
    $worksheet->write(9,  1, '=STDEV(B2:I2)');

    $worksheet->write(10, 0, 'Kurtosis', $format);
    $worksheet->write(10, 1, '=KURT(B2:I2)');



=head2 Example 5 

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




=head1 LIMITATIONS

The following limits are imposed by Excel or the version of the BIFF file that has been implemented:

    Description                          Limit   Source
    -----------------------------------  ------  -------
    Maximum number of chars in a string  255     Excel 5
    Maximum number of columns            256     Excel 5, 97
    Maximum number of rows in Excel 5    16384   Excel 5
    Maximum number of rows in Excel 97   65536   Excel 97


Note: the maximum row reference in a formula is the Excel 5 row limit of 16384.

The minimum file size is 6K due to the OLE overhead. The maximum file size is approximately 7MB (7087104 bytes) of BIFF data. This can be extended by using Takanori Kawai's OLE::Storage_Lite module http://search.cpan.org/search?dist=OLE-Storage_Lite see the C<big.pl> example in the C<examples> directory of the distro.




=head1 REQUIREMENTS

This module requires Perl 5.005 (or later) and Parse::RecDescent: http://search.cpan.org/search?dist=Parse-RecDescent




=head1 PORTABILITY

Spreadsheet::WriteExcel.pm will work on the majority of Windows, UNIX and Macintosh platforms. Specifically, the module will work on any system where perl packs floats in the 64 bit IEEE format. The float must also be in little-endian format but WriteExcel.pm will reverse it as necessary. Thus:

    print join(" ", map { sprintf "%#02x", $_ } unpack("C*", pack "d", 1.2345)), "\n";

should give (or in reverse order):

    0x8d 0x97 0x6e 0x12 0x83 0xc0 0xf3 0x3f

In general, if you don't know whether your system supports a 64 bit IEEE float or not, it probably does. If your system doesn't, WriteExcel will C<croak()> with the message given in the L<DIAGNOSTICS> section. You can check which platforms the module has been tested on at the CPAN testers site: http://testers.cpan.org/search?request=dist&dist=Spreadsheet-WriteExcel




=head1 DIAGNOSTICS

=over 4

=item Filename required in WriteExcel('Filename')

A filename must be given in the constructor.

=item Can't open filename. It may be in use.

The file cannot be opened for writing. The directory that you are writing to  may be protected or the file may be in use by another program.

=item Required floating point format not supported on this platform.

Operating system doesn't support 64 bit IEEE float or it is byte-ordered in a way unknown to WriteExcel.

=item Unable to create tmp files via IO::File->new_tmpfile().

This is a C<-w> warning. You will see it if you are using Spreadsheet::WriteExcel in an environment where temporary files cannot be created, in which case all data will be stored in memory. The warning is for information only: it does not affect execution but it may effect the speed of execution for large files.

=item Maximum file size, 7087104, exceeded.

The current OLE implementation only supports a maximum BIFF file of this size. This limit can be extended, see the L<LIMITATIONS> section.

=item Can't locate Parse/RecDescent.pm in @INC ...

Spreadsheet::WriteExcel requires the Parse::RecDescent module. Download it from CPAN: http://search.cpan.org/search?dist=Parse-RecDescent

=item Couldn't parse formula ...

There are a large number of warnings which relate to badly formed formulas and functions. See the L<FORMULAS AND FUNCTIONS IN EXCEL> section for suggestions on how to avoid these errors.

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

The source code for the Excel plugin of the Gnumeric spreadsheet also contains information relevant to the Excel BIFF format and the OLE container, http://www.gnumeric.org/gnumeric

In addition the source code for OpenOffice is available at http://www.openoffice.org/

An article describing Spreadsheet::WriteExcel and how it works appears in Issue #19 of The Perl Journal, http://www.itknowledge.com/tpj/ It is reproduced, by kind permission, in the C<doc> directory of the distro.


Please note that the provision of this information does not constitute an invitation to start hacking at the BIFF or OLE file formats. There are more interesting ways to waste your time. ;-)




=head1 WRITING EXCEL FILES

Depending on your requirements, background and general sensibilities you may prefer one of the following methods of getting data into Excel:

* CSV, comma separated variables or text. If the file extension is C<csv>, Excel will open and convert this format automatically. Generating a valid CSV file isn't as easy as it seems. Have a look at the DBD::RAM, DBD::CSV and Text::CSV_XS modules.

* DBI with DBD::ADO or DBD::ODBC. Excel files contain an internal index table that allows them to act like a database file. Using one of the standard Perl database modules you can connect to an Excel file as a database.

* Win32::OLE module and office automation. This requires a Windows platform and an installed copy of Excel. This is the most powerful and complete method for interfacing with Excel. See http://velocity.activestate.com/docs/ActivePerl/site/lib/Win32/OLE/TPJ.html , http://velocity.activestate.com/docs/ActivePerl/faq/Windows/ActivePerl-Winfaq12.html and http://velocity.activestate.com/docs/ActivePerl/site/lib/Win32/OLE.html If your main platform is UNIX but you have the resources to set up a separate Win32/MSOffice server, you can convert office documents to text, postscript or PDF using Win32::OLE. For a demonstration of how to do this using Perl see Docserver: http://search.cpan.org/search?mode=module&query=docserver

* HTML tables. This is an easy way of adding formatting.

* XML, the Excel XML and HTML file specification are available from http://msdn.microsoft.com/library/officedev/ofxml2k/ofxml2k.htm




=head1 READING EXCEL FILES

To read data from Excel files try:

* Spreadsheet::ParseExcel. This uses the OLE::Storage-Lite module to extract data from an Excel file. http://search.cpan.org/search?dist=Spreadsheet-ParseExcel

* OLE::Storage, aka LAOLA. This is a Perl interface to OLE file formats. In particular, the distro contains an Excel to HTML converter called Herbert, http://user.cs.tu-berlin.de/~schwartz/pmh/ There is also an open source C/C++ project based on the LAOLA work. Try the Filters Project http://arturo.directmail.org/filtersweb/ and the Excel to HTML converter at the xlHtml Project http://www.xlhtml.org/

* HTML tables. If the files are saved from Excel in a HTML format the data can be accessed using HTML::TableExtract http://search.cpan.org/search?dist=HTML-TableExtract

* DBI with DBD::ADO or DBD::ODBC. See, the section L<WRITING EXCEL FILES>.

* Win32::OLE module and office automation. See, the section L<WRITING EXCEL FILES>.

If you wish to view Excel files on a UNIX/Linux platform check out the excellent Gnumeric spreadsheet application at http://www.gnumeric.org/gnumeric or OpenOffice at http://www.openoffice.org/

If you wish to view Excel files on a Windows platform which doesn't have Excel installed you can use the free Microsoft Excel Viewer http://officeupdate.microsoft.com/2000/downloaddetails/xlviewer.htm




=head1 BUGS

Orange isn't.

Formulas are formulae.

OpenOffice: Numerical formats are not displayed due to some missing records in Spreadsheet::WriteExcel. Someone with a good knowledge of C++, and possibly of German, might help me to track this down in the OpenOffice source. URLs are not displayed as links.

Gnumeric: Some formatting is not displayed correctly. URLs are not displayed as links.

MS Access: The Excel files that are produced by this module are not compatible with MS Access. Use DBI or ODBC instead.

QuickView: If you wish to write files that are fully compatible with QuickView it is necessary to write the cells in a sequential row by row order.

The lack of a portable way of writing a little-endian 64 bit IEEE float.



=head1 TO DO

Read a book, see a film or two.




=head1 ACKNOWLEDGEMENTS

The following people contributed to the debugging and testing of Spreadsheet::WriteExcel:

Arthur@ais, Artur Silveira da Cunha, Cedric Bouvier, CPAN testers, Daniel Gardner, Harold Bamford, James Holmes, Johan Ekenberg, John Wren, Kenneth Stacey, Michael Buschauer, Mike Blazer, Paul J. Falbe, Paul Medynski, Shenyu Zheng, Rich Sorden.

The following people contributed code or examples:

Andrew Benham, Marco Geri, Sam Kington, Takanori Kawai.

Thanks to Damian Conway for the excellent Parse::RecDescent. Thanks to Michael Meeks for Gnumeric.




=head1 AUTHOR

John McNamara jmcnamara@cpan.org


    Lana Turner has collapsed!
    I was trotting along and suddenly
    it started raining and snowing
    and you said it was hailing
    but hailing hits you on the head
    hard so it was really snowing and
    raining and I was in such a hurry
    to meet you but the traffic
    was acting exactly like the sky
    and suddenly I see a headline
    LANA TURNER HAS COLLAPSED!
    there is no snow in Hollywood
    there is no rain in California
    I have been to lots of parties
    and acted perfectly disgraceful
    but I never actually collapsed
    oh Lana Turner we love you get up
        --Frank O'Hara


=head1 COPYRIGHT

© MM-MMI, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

