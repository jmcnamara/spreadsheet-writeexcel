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

$VERSION = '0.32'; #  Weller.



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

This document refers to version 0.32 of Spreadsheet::WriteExcel, released May 17, 2001.




=head1 SYNOPSIS

To write a string, a formatted string, a number and a formula to the first worksheet in an Excel workbook called perl.xls:

    use Spreadsheet::WriteExcel;

    # Create a new Excel workbook
    my $workbook = Spreadsheet::WriteExcel->new("perl.xls");

    # Add a worksheet
    $worksheet = $workbook->addworksheet();

    #  Add and define a format
    $format = $workbook->addformat(); # Add a format
    $format->set_bold();
    $format->set_color('red');
    $format->set_align('center');

    # Write a formatted and unformatted string, row and column notation.
    $col = $row = 0;
    $worksheet->write($row, $col, "Hi Excel!", $format);
    $worksheet->write(1,    $col, "Hi Excel!");

    # Write a number and a formula using A1 notation
    $worksheet->write('A3', 1.2345);
    $worksheet->write('A4', '=SIN(PI()/4)');




=head1 DESCRIPTION

The Spreadsheet::WriteExcel module can be used to create a cross-platform Excel binary file. Multiple worksheets can be added to a workbook and formatting can be applied to cells. Text, numbers, formulas and hyperlinks can be written to the cells.

The Excel file produced by this module is compatible with Excel 5, 95, 97 and 2000.

The module will work on the majority of Windows, UNIX and Macintosh platforms. Generated files are also compatible with the Linux/UNIX spreadsheet applications OpenOffice, Gnumeric and XESS. The generated files are not compatible with MS Access.




=head1 WORKBOOK METHODS

The Spreadsheet::WriteExcel module provides an object oriented interface to a new Excel workbook. The following methods are available through a new workbook.

If you are unfamiliar with object oriented interfaces or the way that they are implemented in Perl have a look at C<perlobj> and C<perltoot> in the main Perl documentation.


=head2 new()

A new Excel workbook is created using the C<new()> constructor which accepts either a filename or a filehandle as a parameter. The following example creates a new Excel file based on a filename:

    my $workbook  = Spreadsheet::WriteExcel->new('filename.xls');
    my $worksheet = $workbook->addworksheet();
    $worksheet->write(0, 0,  "Hi Excel!");

Here are some other examples of using C<new()> with filenames:

    my $workbook1 = Spreadsheet::WriteExcel->new($filename);
    my $workbook2 = Spreadsheet::WriteExcel->new("/tmp/filename.xls");
    my $workbook3 = Spreadsheet::WriteExcel->new("c:\\tmp\\filename.xls");
    my $workbook4 = Spreadsheet::WriteExcel->new('c:\tmp\filename.xls');

The last two examples demonstrates how to create a file on DOS or Windows where it is necessary to either escape the directory separator C<\> or to use single quotes to ensure that it isn't interpolated. For more information  see C<perlfaq5: Why can't I use "C:\temp\foo" in DOS paths?>.

The C<new()> constructor returns a Spreadsheet::WriteExcel object that you can use to add worksheets and store data. It should be noted that although C<my> is not specifically required it defines the scope of the new workbook variable and, in the majority of cases, ensures that the workbook is closed properly without explicitly calling the C<close()> method.

You can also pass a valid filehandle to the C<new()> constructor. For example in a CGI program you could do something like this:

    binmode(STDOUT);
    my $workbook  = Spreadsheet::WriteExcel->new(\*STDOUT);

The requirement for C<binmode()> is explained below.

For CGI programs you can also use the special Perl filename C<'-'> which will redirect the output to STDOUT:

    my $workbook  = Spreadsheet::WriteExcel->new('-');

See also, the C<cgi.pl> program in the C<examples> directory of the distro. However, this special case will not work in C<mod_perl> programs where you will have to do something like the following:

    tie *XLS, 'Apache';
    binmode(XLS);
    my $workbook  = Spreadsheet::WriteExcel->new(\*XLS);

Filehandles can also be useful if you want to stream an Excel file over a socket or if you want to store an Excel file in a tied scalar. For some examples of using filehandles with Spreadsheet::WriteExcel see the C<filehandle.pl> program in the C<examples> directory of the distro.

Note about the requirement for C<binmode()>: An Excel file is comprised of binary data. Therefore, if you are using a filehandle you should ensure that you C<binmode()> it prior to passing it to C<new()>.You should do this regardless of whether your platform requires it or not. For more information about C<binmode()> see C<perlfunc> and C<perlopentut> in the main Perl documentation. It is equally important to note that you do not need to C<binmode()> a filename. In fact it would cause an error. Spreadsheet::WriteExcel performs the C<binmode()> internally when it converts the filename to a filehandle.




=head2 close()

The C<close()> method can be used to explicitly close an Excel file.

    $workbook->close();

An explicit C<close()> is required if the file must be closed prior to performing some external action on it such as copying or reading its size.

In addition, C<close()> may be required if the scope of the Workbook, Worksheet or Format objects cannot be determined by perl. Situations where this can occur are:

=over

=item * If C<my()> was not used to declare the scope of a workbook variable created using C<new()>.

=item * If the C<addworksheet()> or C<addformat()> methods are called in subroutines.

=back

The reason for this is that Spreadsheet::WriteExcel relies on Perl's C<DESTROY> mechanism to trigger destructor methods in a specific sequence. This will not happen if the scope of the variables cannot be determined.


In general, if you create a file with a size of 0 bytes you need to call C<close()>.




=head2 addworksheet($sheetname)

At least one worksheet should be added to a new workbook. A worksheet is used to write data into cells:

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




=head2 Cell notation

Spreadsheet::WriteExcel supports two forms of notation to designate the position of cells: Row-column notation and A1 notation.

Row-column notation uses a zero based index for both row and column while A1 notation uses the standard Excel alphanumeric sequence of column letter and 1-based row. For example:

    (0, 0)      # The top left cell in row-column notation.
    ('A1')      # The top left cell in A1 notation.

    (1999, 29)  # Row-column notation.
    ('AD2000')  # The same cell in A1 notation.

Row-column notation is useful if you are referring to cells programmatically:

    for my $i (0 .. 9) {
        $worksheet->write($i, 0, 'Hello'); # Cells A1 to A10
    }

A1 notation is useful for setting up a worksheet manually and for working with formulas:

    $worksheet->write('H1', 200);
    $worksheet->write('H2', '=H7+1');

In the C<examples> directory of the distro there is a program called C<convertA1.pl> which contains helper functions for dealing with A1 notation.

For simplicity, the parameter lists for the worksheet method calls in the following sections are given in terms of row-column notation. In all cases it is also possible to use A1 notation.




=head2 write($row, $column, $token, $format)

The C<write()> method is a general alias for one of several methods of writing to a cell in Excel. C<write()> calls one of the following methods depending on the value of C<$token>:

C<write_number()> if C<$token> is a number based on the following regex: C<$token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/>.

C<write_blank()> if C<$token> is a blank string: C<""> or C<''>.

C<write_url()> if C<$token> is a http, ftp or mailto URL based on the following regexes: C<$token =~ m|^[fh]tt?p://|> or  C<$token =~ m|^mailto:|>.

C<write_formula()> if the first character of C<$token> is C<"=">.

C<write_string()> if none of the previous conditions apply.

Here are some examples in both row-column and A1 notation:

    $worksheet->write(0, 0, "Hello"                 );  # write_string()
    $worksheet->write('A2', 'One'                   );  # write_string()
    $worksheet->write(2, 0,  2                      );  # write_number()
    $worksheet->write('A4',  3.00001                );  # write_number()
    $worksheet->write(4, 0,  ""                     );  # write_blank()
    $worksheet->write('A6',  ''                     );  # write_blank()
    $worksheet->write(6, 0,  'http://www.perl.com/' );  # write_url()
    $worksheet->write('A8',  'ftp://ftp.cpan.org/'  );  # write_url()
    $worksheet->write(8, 0,  '=A3 + 3*A4'           );  # write_formula()
    $worksheet->write('A10', '=SIN(PI()/4)'         );  # write_formula()

The C<$format> argument is optional. It should be a valid Format object, see L<FORMAT METHODS>:

    my $format = $workbook->addformat();
    $format->set_bold();
    $format->set_color('red');
    $format->set_align('center');

    $worksheet->write(4, 0, "Hello", $format ); # Formatted string


The C<write> methods return:

    0 for success.
   -1 for insufficient number of arguments.
   -2 for row or column out of bounds.
   -3 for string too long.




=head2 write_number($row, $column, $number, $format)

Write an integer or a float to the cell specified by C<$row> and C<$column>:

    $worksheet->write_number(0, 0,  1     );
    $worksheet->write_number('A2',  2.3451);

See the note about L<Cell notation>. The C<$format> argument is optional.




=head2 write_string($row, $column, $string, $format)

Write a string to the cell specified by C<$row> and C<$column>:

    $worksheet->write_string(0, 0, "Your text here" );
    $worksheet->write_string('A2', "or here" );

See the note about L<Cell notation>. The maximum string size is 255 characters. The C<$format> argument is optional.




=head2 write_formula($row, $column, $formula, $format)

Write a formula or function to the cell specified by C<$row> and C<$column>:

    $worksheet->write_formula(0, 0, '=$B$3 + B4'  );
    $worksheet->write_formula(1, 0, '=SIN(PI()/4)');
    $worksheet->write_formula(2, 0, '=SUM(B1:B5)' );
    $worksheet->write_formula('A4', '=IF(A3>1,"Yes", "No")'   );
    $worksheet->write_formula('A5', '=AVERAGE(1, 2, 3, 4)'    );
    $worksheet->write_formula('A6', '=DATEVALUE("1-Jan-2001")');

See the note about L<Cell notation>. For more information about writing Excel formulas see L<FORMULAS AND FUNCTIONS IN EXCEL>




=head2 write_blank($row, $column, $format)

Write a blank cell specified by C<$row> and C<$column>:

    $worksheet->write_blank(0, 0, $format);

See the note about L<Cell notation>. This method is useful for adding formatting to a cell which doesn't contain a string or number value.




=head2 write_url($row, $col, $url, $string, $format)

Write a hyperlink to a URL in the cell specified by C<$row> and C<$column>. The hyperlink is comprised of two elements: the visible label and the invisible link. The visible label is the same as the link unless an alternative string is specified. The alternative C<$string> and the C<$format> are optional.

    $worksheet->write_url(0, 0, 'http://www.perl.com/'                );
    $worksheet->write_url(1, 0, 'http://www.perl.com/', 'Perl home'   );
    $worksheet->write_url(2, 0, 'http://www.perl.com/', undef, $format);
    $worksheet->write_url(3, 0, 'mailto:jmcnamara@cpan.org'           );

The label is written using the C<write_string()> method. Therefore the 255 characters string limit applies to the label: the URL can be any length. Use C<undef> if you wish to specify a format without specifying an alternative string.See the note about L<Cell notation>.

Note: Hyperlinks are not available in Excel 5. They will appear as a string only.




=head2 activate()

The C<activate()> method is used to specify which worksheet is initially visible in a multi-sheet workbook:

    $worksheet1 = $workbook->addworksheet('To');
    $worksheet2 = $workbook->addworksheet('the');
    $worksheet3 = $workbook->addworksheet('wind');

    $worksheet3->activate();

This is similar to the Excel VBA activate method. More than one worksheet can be selected via the C<select()> method, however only one worksheet can be active. The default value is the first worksheet.




=head2 select()

The C<select()> method is used to indicate that a worksheet is selected in a multi-sheet workbook:

    $worksheet1->activate();
    $worksheet2->select();
    $worksheet3->select();

A selected worksheet has its tab highlighted. Selecting worksheets is a way of grouping them together so that, for example, several worksheets could be printed in one go.. A worksheet that has been activated via the C<activate()> method will also appear as a selected. You probably won't need to use the C<select()> method very often.




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

This method can be used to specify which cell or cells are selected in a worksheet. The most common requirement is to select a single cell, in which case C<$last_row> and C<$last_col> are not required. The active cell within a selected range is determined by the order in which C<$first> and C<$last> are specified. It is also possible to specify a cell or a range using A1 notation. See the note about L<Cell notation>.

Examples:

    $worksheet1->set_selection(3, 3);       # 1. Cell D4.
    $worksheet2->set_selection(3, 3, 6, 6); # 2. Cells D4 to G7.
    $worksheet3->set_selection(6, 6, 3, 3); # 3. Cells G7 to D4.
    $worksheet4->set_selection('D4');       # Same as 1.
    $worksheet5->set_selection('D4:G7');    # Same as 2.
    $worksheet6->set_selection('G7:D4');    # Same as 3.

The default cell is (0, 0), 'A1'.




=head2 set_row($row, $height, $format);

This method can be used to specify the height of a row. The C<$format> argument is optional, for additional information, see L<FORMAT METHODS>.

    $worksheet->set_row(0, 20); # Row 1 height set to 20

If you wish to set the format without changing the height you can pass C<undef> as the height parameter:

    $worksheet->set_row(0, undef, $format);


=head2 set_column($first_col, $last_col, $width, $format);

This method can be used to specify the width of a single column or a range of columns. If the method is applied to a single column the value of C<$first_col> and C<$last_col> should be the same. It is also possible to specify a column range using the form of A1 notation used for columns. See the note about L<Cell notation>.

Examples:

    $worksheet->set_column(0, 0,  20); # Column  A   width set to 20
    $worksheet->set_column(1, 3,  30); # Columns B-D width set to 30
    $worksheet->set_column('E:E', 20); # Column  E   width set to 20
    $worksheet->set_column('F:H', 30); # Columns F-H width set to 30

The width corresponds to the column width value that is specified in Excel. It is approximately equal to the length of a string in the default font of Arial 10. The C<$format> argument is optional, for additional information, see L<FORMAT METHODS>.

If you wish to set the format without changing the width you can pass C<undef> as the width parameter:

    $worksheet->set_column(0, 0, undef, $format);

Unfortunately, there is no way to specify "AutoFit" for a column in the Excel file format. This feature is only available at runtime from within Excel.



=head2 freeze_panes($row, $col, $top_row, $left_col);

This method can be used to divide a worksheet into horizontal or vertical regions known as panes and to also "freeze" these panes so that the splitter bars are not visible. This is the the same as the "Window->Freeze Panes" menu command in Excel

The parameters C<$row> and C<$col> are used to specify the location of the split. It should be noted that the split is specified at the top or left of a cell and that the method uses zero based indexing. Therefore to freeze the first row of a worksheet it is necessary to specify the split at row 2 (which is 1 as the zero-based index). This might lead you to think that you are using a 1 based index but this is not the case.

You can set one of the C<$row> and C<$col> parameters as zero if you do not want either a vertical or horizontal split. 

Examples:

    $worksheet->freeze_panes(1, 0); # Freeze the first row
    $worksheet->freeze_panes('A2'); # Same using A1 notation
    $worksheet->freeze_panes(0, 1); # Freeze the first column
    $worksheet->freeze_panes('B1'); # Same using A1 notation
    $worksheet->freeze_panes(1, 2); # Freeze the first row and first 2 columns
    $worksheet->freeze_panes('C2'); # Same using A1 notation

The parameters C<$top_row> and C<$left_col> are optional. They are used to specify the top-most or left-most visible row or column in the scrolling region of the panes. For example to freeze the first row and to have the scrolling region begin at row twenty:

    $worksheet->freeze_panes(1, 0, 20, 0);

You cannot use A1 notation for the C<$top_row> and C<$left_col> parameters.
    

See also the C<panes.pl> program in the examples directory of the distribution.




=head2 thaw_panes($y, $x, $top_row, $left_col);

This method can be used to divide a worksheet into horizontal or vertical regions known as panes. This method is different from the C<freeze_panes()> method in that the splits between the panes will be visible to the user and each pane will have its own scroll bars.

The parameters C<$y> and C<$x> are used to specify the vertical and horizontal position of the split. The units for C<$y> and C<$x> are the same as those used by Excel to specify row height and column width. However, the vertical and horizontal units are different from each other. Therefore you must specify the C<$y> and C<$x> parameters in terms of the row heights and column widths that you have set or the default values which are C<12.75> for a row and  C<8.43> for a column.

You can set one of the C<$y> and C<$x> parameters as zero if you do not want either a vertical or horizontal split. The parameters C<$top_row> and C<$left_col> are optional. They are used to specify the top-most or left-most visible row or column in the bottom-right pane.

Example:

    $worksheet->thaw_panes(12.75, 0,    1, 0); # First row
    $worksheet->thaw_panes(0,     8.43, 0, 1); # First column
    $worksheet->thaw_panes(12.75, 8.43, 1, 1); # First row and column

You cannot use A1 notation with this method.

See also the C<freeze_panes()> method and the C<panes.pl> program in the examples directory of the distribution.




=head1 PAGE SETUP METHODS

Page setup methods affect the way that a worksheet will look when it is printed. They control features such as page headers and footers and margins. These methods are really just standard worksheet methods. They are documented here in a separate section for the sake of clarity.

Not all of Excel's page setup options are available in this release but they will be added in subsequent releases.

A common requirement when working with Spreadsheet::WriteExcel is to apply the same page setup features to all of the worksheets in a workbook. To do this you can use the the C<worksheets()> method of the C<workbook> class to access the array of worksheets in a workbook:

    foreach $worksheet (@{$workbook->worksheets()}) {
       $worksheet->$worksheet->set_orientation(0);
    }




=head2 set_landscape()

This method is used to set the orientation of a worksheet's printed page to landscape:

    $worksheet->set_orientation(); # Landscape mode




=head2 set_portrait()

This method is used to set the orientation of a worksheet's printed page to portrait. The default worksheet orientation is portrait, so you generally won't need to call this method.

    $worksheet->set_portrait(); # Portrait mode



=head2 set_paper($index)

This method is used to set the paper format for the printed output of a worksheet. The following paper styles are available:

    Index   Paper format            Paper size
    =====   ============            ==========
      0     Printer default         -
      1     Letter                  8 1/2 x 11 in
      2     Letter Small            8 1/2 x 11 in
      3     Tabloid                 11 x 17 in
      4     Ledger                  17 x 11 in
      5     Legal                   8 1/2 x 14 in
      6     Statement               5 1/2 x 8 1/2 in
      7     Executive               7 1/4 x 10 1/2 in
      8     A3                      297 x 420 mm
      9     A4                      210 x 297 mm
     10     A4 Small                210 x 297 mm
     11     A5                      148 x 210 mm
     12     B4                      250 x 354 mm
     13     B5                      182 x 257 mm
     14     Folio                   8 1/2 x 13 in
     15     Quarto                  215 x 275 mm
     16     -                       10x14 in
     17     -                       11x17 in
     18     Note                    8 1/2 x 11 in
     19     Envelope  9             3 7/8 x 8 7/8
     20     Envelope 10             4 1/8 x 9 1/2
     21     Envelope 11             4 1/2 x 10 3/8
     22     Envelope 12             4 3/4 x 11
     23     Envelope 14             5 x 11 1/2
     24     C size sheet            -
     25     D size sheet            -
     26     E size sheet            -
     27     Envelope DL             110 x 220 mm
     28     Envelope C3             324 x 458 mm
     29     Envelope C4             229 x 324 mm
     30     Envelope C5             162 x 229 mm
     31     Envelope C6             114 x 162 mm
     32     Envelope C65            114 x 229 mm
     33     Envelope B4             250 x 353 mm
     34     Envelope B5             176 x 250 mm
     35     Envelope B6             176 x 125 mm
     36     Envelope                110 x 230 mm
     37     Monarch                 3.875 x 7.5 in
     38     Envelope                3 5/8 x 6 1/2 in
     39     Fanfold                 14 7/8 x 11 in
     40     German Std Fanfold      8 1/2 x 12 in
     41     German Legal Fanfold    8 1/2 x 13 in


Note, it is likely that not all of these paper types will be available to the end user since it will depend on the paper formats that the user's printer supports. Therefore, it is best to stick to standard paper types.

If you do not specify a paper type the worksheet will print using the printer's default paper. 

    $worksheet->set_paper(1); # US Letter
    $worksheet->set_paper(9); # A4



=head2 center_horizontally()

Center the worksheet data horizontally between the margins on the printed page:

    $worksheet->center_horizontally();




=head2 center_vertically()

Center the worksheet data vertically between the margins on the printed page:

    $worksheet->center_vertically();




=head2 set_margins($inches)

There are several methods available for setting the worksheet margins on the printed page:

    set_margins()        # Set all margins to the same value
    set_margins_LR()     # Set left and right margins to the same value
    set_margins_TB()     # Set top and bottom margins to the same value
    set_margin_left();   # Set left margin
    set_margin_right();  # Set right margin
    set_margin_top();    # Set top margin
    set_margin_bottom(); # Set bottom margin

All of these methods take a distance in inches as a parameter. Note: 1 inch = 25.4mm. ;-) The default left and right margin is 0.75 inch. The default top and bottom margin is 1.00 inch. 



=head2 set_header($string, $margin)

Headers and footers are generated using a C<$string> which is a combination of plain text and control characters. The C<$margin> parameter is optional.

The available control character are:

    Control              Type             Description
    =======              ====             ===========
    &L                   Justification    Left
    &C                                    Center
    &R                                    Right
    
    &P                   Information      Page number
    &N                                    Total number of pages
    &D                                    Date
    &T                                    Time
    &F                                    File name
    &A                                    Worksheet name
    
    &fontsize            Font             Font size
    &"fontname,style"                     Font name and style
    
    

Text in headers and footers can be justified (aligned) to the left, center and right by prefixing the text with the control characters C<&L>, C<&C> and C<&R>.

For example (with ASCII art representation of the results):

    $worksheet->set_header('&LHello');
    
     ---------------------------------------------------------------
    |                                                               |
    | Hello                                                         |
    |                                                               |
    
    
    $worksheet->set_header('&CHello');
    
     ---------------------------------------------------------------
    |                                                               |
    |                          Hello                                |
    |                                                               |
    
    
    $worksheet->set_header('&RHello');
    
     ---------------------------------------------------------------
    |                                                               |
    |                                                         Hello |
    |                                                               |
    

If you do not specify any justification the text will be centered:

    $worksheet->set_header('Hello');
    
     ---------------------------------------------------------------
    |                                                               |
    |                          Hello                                |
    |                                                               |
    

You can also have text in each of the justification regions:

    $worksheet->set_header('&LCiao&CBello&RCielo');
    
     ---------------------------------------------------------------
    |                                                               |
    | Ciao                     Bello                          Cielo |
    |                                                               |
    

The information control characters act as variables that Excel will update as the workbook or worksheet changes. Times and dates are in the users default format:

    $worksheet->set_header('Page &P of &N');
    
     ---------------------------------------------------------------
    |                                                               |
    |                        Page 1 of 6                            |
    |                                                               |
    
    
    $worksheet->set_header('Updated at &T');
    
     ---------------------------------------------------------------
    |                                                               |
    |                    Updated at 12:30 PM                        |
    |                                                               |
    


You can specify the font size of a section of the text by prefixing it with the control character C<&n> where C<n> is the font size:

    $worksheet->set_header('&20Big Hello');
    $worksheet->set_header('&6Small Hello');

You can specify the font of a section of the text by prefixing it with the control sequence C<&"fontname,style"> where C<fontname> is a font name such as "Courier New" or "Times New Roman" and C<style> is one of the standard Windows font descriptions: "Regular", "Italic", "Bold" or "Bold Italic":

    $worksheet1->set_header('&"Courier New,Italic"Hello');
    $worksheet2->set_header('&"Courier New,Bold Italic"Hello');
    $worksheet3->set_header('&"Times New Roman,Regular"Hello');

It is possible to combine all of these features together to create sophisticated headers and footers. As an aid to setting up complicated headers and footers you can record a page setup as a macro in Excel and look at the format strings that VBA produces. Remember however that VBA uses a double double quote C<""> to indicate a single double quote. For the last example above the equivalent VBA code looks like this:

    .LeftHeader   = ""
    .CenterHeader = "&""Times New Roman,Regular""Hello"
    .RightHeader  = ""

The header or footer string must be less than 255 characters.

As stated above the margin parameter is optional. As with the other margins the value should be in inches. The default header and footer margin is 0.50 inch. The header and footer margin size can be set as follows:

    $worksheet->set_header('&CHello', 0.75);

Note, the header and footer margins are independent of the top and bottom margins.



=head2 set_footer()

See the C<set_header()> section above.




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
    * You cannot use arrays constants, i.e. {1;2;3}, in functions.
    * You cannot use embedded double quotes in strings.
    * Whitespace is not preserved around operators.

However, these constraints will be removed in future versions. They are here because of a trade-off between features and time.

The following is a brief introduction to formulas and functions in Excel and Spreadsheet::WriteExcel.

A formula is a string that begins with an equal sign:

    '=A1+B1'
    '=AVERAGE(1, 2, 3)'

The formula can contain numbers, strings, boolean values, cell references, cell ranges and functions. Formulas should be written as they appear in Excel, that is cells and functions must be in uppercase.

Cells in Excel are referenced using the A1 notation system where the column is designated by a letter and the row by a number. Columns range from A to IV i.e. 0 to 255, rows range from 1 to 16384. In the C<examples> directory of the distro there is program called C<convertA1.pl> which contains functions to help you work with this system.

The Excel C<$> notation in cell references is also supported. This allows you to specify whether a row or column is relative or absolute. This only has an effect if the cell is copied. The following examples show relative and absolute values.

    '=A1'   # Column and row are relative
    '=$A1'  # Column is absolute and row is relative
    '=A$1'  # Column is relative and row is absolute
    '=$A$1' # Column and row are absolute

Formulas can also refer to cells in other worksheets of the current workbook. For example:

    '=Sheet2!A1'
    '=Sheet2!A1:A5'
    '=Sheet2:Sheet3!A1'
    '=Sheet2:Sheet3!A1:A5'
    q{='Test Data'!A1}
    q{='Test Data1:Test Data2'!A1}

The sheet reference and and the cell reference are separated by  C<!> the exclamation mark symbol. If worksheet names contain spaces then Excel requires them to be enclosed in single quotes as in the last two examples above. In this case you will have to use the quote operator C<q{}> to protect the quotes. See C<perlop> in the main Perl documentation. Only valid sheet names that have been added using the C<addworksheet()> method can be used in formulas. You cannot reference external workbooks.


The following table lists the operators that are available in Excel's formulas. The majority of the operators are the same as Perl's. Differences are indicated.

    Arithmetic operators:
    =====================
    Operator  Meaning                   Example
       +      Addition                  1+2
       -      Subtraction               2-1
       *      Multiplication            2*3
       /      Division                  1/4
       ^      Exponentiation            2^3      # Equivalent to **
       -      Unary minus               -(1+2)   # Not yet supported
       %      Percent (Not modulus)     13%      # Not supported, [1]


    Comparison operators:
    =====================
    Operator  Meaning                   Example
        =     Equal to                  A1 =  B1 # Equivalent to ==
        <>    Not equal to              A1 <> B1 # Equivalent to !=
        >     Greater than              A1 >  B1
        <     Less than                 A1 <  B1
        >=    Greater than or equal to  A1 >= B1
        <=    Less than or equal to     A1 <= B1


    String operator:
    ================
    Operator  Meaning                   Example
        &     Concatenation             "Hello " & "World!" # [2]


    Reference operators:
    ====================
    Operator  Meaning                   Example
        :     Range operator            A1:A4               # [3]
        ,     Union operator            SUM(1, 2+2, B3)     # [4]


    Notes:
    [1]: You can get a percentage with formatting and modulus with MOD().
    [2]: Equivalent to ("Hello " . "World!") in Perl.
    [3]: This range is equivalent to cells A1, A2, A3 and A4.
    [4]: The comma behaves like the list separator in Perl.

The range and comma operators can have different symbols in non-English versions of Excel. These will be supported in a later version of Spreadsheet::WriteExcel.

The following table lists all of the core functions supported by Excel 5 and Spreadsheet::WriteExcel. Any additional functions that are available through the "Analysis ToolPak" or other add-ins are not supported. These functions have all been tested to verify that they work.

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

You can also modify the module to support function names in the following languages: German, French, Spanish, Portuguese, Dutch, Finnish, Italian and Swedish. See the C<function_locale.pl> program in the C<examples> directory of the distro.

For a general introduction to Excel's formulas and an explanation of the syntax of the function refer to the Excel help files or the following links: http://msdn.microsoft.com/library/default.asp?URL=/library/officedev/office97/s88f2.htm and http://msdn.microsoft.com/library/default.asp?URL=/library/officedev/office97/s992f.htm


If your formula doesn't work in Spreadsheet::WriteExcel try the following:

    1. Verify that the formula works in Excel (or Gnumeric or OpenOffice).
    2. Ensure that it isn't on the TODO list at the start of this section.
    3. Ensure that cell references and formula names are in uppercase.
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

The minimum file size is 6K due to the OLE overhead. The maximum file size is approximately 7MB (7087104 bytes) of BIFF data. This can be extended by using Takanori Kawai's OLE::Storage_Lite module http://search.cpan.org/search?dist=OLE-Storage_Lite see the C<bigfile.pl> example in the C<examples> directory of the distro.




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

This is a C<-w> warning. You will see it if you are using Spreadsheet::WriteExcel in an environment where temporary files cannot be created, in which case all data will be stored in memory. The warning is for information only: it does not affect execution but it may affect the speed of execution for large files.

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

For a open source implementation of the OLE library see the 'cole' library at http://atena.com/libole2.php

The source code for the Excel plugin of the Gnumeric spreadsheet also contains information relevant to the Excel BIFF format and the OLE container, http://www.gnumeric.org/gnumeric

In addition the source code for OpenOffice is available at http://www.openoffice.org/

An article describing Spreadsheet::WriteExcel and how it works appears in Issue #19 of The Perl Journal, http://www.tpj.com/ It is reproduced, by kind permission, in the C<doc> directory of the distro.


Please note that the provision of this information does not constitute an invitation to start hacking at the BIFF or OLE file formats. There are more interesting ways to waste your time. ;-)




=head1 WRITING EXCEL FILES

Depending on your requirements, background and general sensibilities you may prefer one of the following methods of getting data into Excel:

* CSV, comma separated variables or text. If the file extension is C<csv>, Excel will open and convert this format automatically. Generating a valid CSV file isn't as easy as it seems. Have a look at the DBD::RAM, DBD::CSV and Text::CSV_XS modules.

* DBI with DBD::ADO or DBD::ODBC. Excel files contain an internal index table that allows them to act like a database file. Using one of the standard Perl database modules you can connect to an Excel file as a database.

* DBI::Excel, you can also access Spreadsheet::WriteExcel using the standard DBI interface via Kawai Takanori's DBI::Excel module http://search.cpan.org/search?dist=DBI-Excel.

* Win32::OLE module and office automation. This requires a Windows platform and an installed copy of Excel. This is the most powerful and complete method for interfacing with Excel. See http://www.activestate.com/ASPN/Reference/Products/ActivePerl-5.6/faq/Windows/ActivePerl-Winfaq12.html and http://www.activestate.com/ASPN/Reference/Products/ActivePerl-5.6/site/lib/Win32/OLE.html If your main platform is UNIX but you have the resources to set up a separate Win32/MSOffice server, you can convert office documents to text, postscript or PDF using Win32::OLE. For a demonstration of how to do this using Perl see Docserver: http://search.cpan.org/search?mode=module&query=docserver

* HTML tables. This is an easy way of adding formatting.

* XML, the Excel XML and HTML file specification are available from http://msdn.microsoft.com/library/officedev/ofxml2k/ofxml2k.htm




=head1 READING EXCEL FILES

To read data from Excel files try:

* Spreadsheet::ParseExcel. This uses the OLE::Storage-Lite module to extract data from an Excel file. http://search.cpan.org/search?dist=Spreadsheet-ParseExcel

* OLE::Storage, aka LAOLA. This is a Perl interface to OLE file formats. In particular, the distro contains an Excel to HTML converter called Herbert, http://user.cs.tu-berlin.de/~schwartz/pmh/ This has been superseded by the Spreadsheet::ParseExcel module. There is also an open source C/C++ project based on the LAOLA work. Try the Filters Project http://atena.com/libole2.php and the Excel to HTML converter at the xlHtml Project http://www.xlhtml.org/

* HTML tables. If the files are saved from Excel in a HTML format the data can be accessed using HTML::TableExtract http://search.cpan.org/search?dist=HTML-TableExtract

* DBI with DBD::ADO or DBD::ODBC. See, the section L<WRITING EXCEL FILES>.

* DBI::Excel, you can also access Spreadsheet::ParseExcel using the standard DBI interface via Kawai Takanori's DBI::Excel module http://search.cpan.org/search?dist=DBI-Excel.

* XML::Excel converts Excel files to XML using Spreadsheet::ParseExcel http://search.cpan.org/search?dist=XML-Excel. 

* Win32::OLE module and office automation. See, the section L<WRITING EXCEL FILES>.

If you wish to view Excel files on a UNIX/Linux platform check out the excellent Gnumeric spreadsheet application at http://www.gnome.org/projects/gnumeric/ or OpenOffice at http://www.openoffice.org/

If you wish to view Excel files on a Windows platform which doesn't have Excel installed you can use the free Microsoft Excel Viewer http://officeupdate.microsoft.com/2000/downloaddetails/xlviewer.htm




=head1 BUGS

Orange isn't.

Formulas are formulae.

Spreadsheet::ParseExcel: All formulas created by Spreadsheet::WriteExcel are read as having a value of zero. This is because Spreadsheet::WriteExcel only stores the formula and not the calculated result.

OpenOffice: Numerical formats are not displayed due to some missing records in Spreadsheet::WriteExcel. Someone with a good knowledge of C++, and possibly of German, might help me to track this down in the OpenOffice source. URLs are not displayed as links.

Gnumeric: Some formatting is not displayed correctly. URLs are not displayed as links.

MS Access: The Excel files that are produced by this module are not compatible with MS Access. Use DBI or ODBC instead.

QuickView: If you wish to write files that are fully compatible with QuickView it is necessary to write the cells in a sequential row by row order.

The lack of a portable way of writing a little-endian 64 bit IEEE float.



=head1 TO DO

The roadmap is as follows:

=over

=item * More page set-up features. The aim is to support all the page setup features that are available through Excel's page setup dialog. Also vertical and horizontal page breaks will be added.

=item * Add cell protection and formula hiding.

=item * Add handling for non-English formula syntax and numbers in different locales.

=item * Move to Excel97/2000 format as standard. This will allow strings greater than 255 characters (who thought up that limit) and hopefully Unicode. The Excel 5 format will be optional.

=back

Also, here are some of the most requested features that probably won't get added:

=over

=item * Graphs. The format is documented but it would require too much work to implement. It would also require too much work to design a useable interface to the hundreds of features in an Excel graph. So that's two too much works.

=item * Macros. This would solve the previous problem neatly. However, the format of Excel macros isn't documented.

=item * Some feature that you really need. ;-)

=back

If there is some feature of an Excel file that you really, really need then you should use Win32::OLE with Excel on Windows. If you are on Unix, then set up a Windows server and use SOAP or CORBA or get your clients to use Gnumeric, it's better than Excel anyway.


=head1 ACKNOWLEDGEMENTS

The following people contributed to the debugging and testing of Spreadsheet::WriteExcel:

Alexander Farber, Arthur@ais, Artur Silveira da Cunha, Borgar Olsen, Cedric Bouvier, CPAN testers, Daniel Berger, Daniel Gardner, Harold Bamford, James Holmes, Johan Ekenberg, J.C. Wren, Kenneth Stacey, Michael Buschauer, Mike Blazer, Paul J. Falbe, Paul Medynski, Peter Dintelmann, Rich Sorden, Shane Ashby, Shenyu Zheng.

The following people contributed code or examples:

Andrew Benham, Bill Young, Marco Geri, Sam Kington, Takanori Kawai, Tom O'Sullivan.

Additional thanks to Takanori Kawai for translating the documentation into Japanese: http://member.nifty.ne.jp/hippo2000/perltips/Spreadsheet/WriteExcel.htm

Thanks to Damian Conway for the excellent Parse::RecDescent. Thanks to Michael Meeks for his work on Gnumeric.




=head1 AUTHOR

John McNamara jmcnamara@cpan.org


    Waking up from bad dreams and smoking cigarettes
    Cuddling a warm girl and smelling stale perfume
    A hot summer's day, and sticky black tarmac
    Feeding ducks in the park, and wishing you were far away
    
    That's entertainment
    
    Two lovers kissing amongst the scream of midnight
    Two lovers missing the tranquility of solitude
    Getting a cab and travelling on buses
    Reading the graffiti about slash-seat affairs
    
    That's entertainment.
    
        --Paul Weller


=head1 COPYRIGHT

© MM-MMI, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

