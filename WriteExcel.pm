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

use Exporter;

use strict;
use Spreadsheet::WriteExcel::Workbook;



use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::WriteExcel::Workbook Exporter);

$VERSION = '0.35'; # The Man with Two Brains.



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
    
    # Check for file creation failures before re-blessing
    bless  $self, $class if defined $self;
    
    return $self;
}


1;


__END__



=head1 NAME

Spreadsheet::WriteExcel - Write to a cross-platform Excel binary file.




=head1 VERSION

This document refers to version 0.35 of Spreadsheet::WriteExcel, released March 18, 2002.




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

The Spreadsheet::WriteExcel module can be used to create a cross-platform Excel binary file. Multiple worksheets can be added to a workbook and formatting can be applied to cells. Text, numbers, formulas, hyperlinks and images can be written to the cells.

The Excel file produced by this module is compatible with Excel 5, 95, 97, 2000 and 2002.

The module will work on the majority of Windows, UNIX and Macintosh platforms. Generated files are also compatible with the Linux/UNIX spreadsheet applications Gnumeric and OpenOffice.

This module cannot be used to write to an existing Excel file.




=head1 QUICK START

Spreadsheet::WriteExcel tries to provide an interface to as many of Excel's features as possible. As a result there is a lot of documentation to accompany the interface and it can be difficult at first glance to see what it important and what is not. So for those of you who prefer to assemble Ikea furniture first and then read the instructions, here are three easy steps: 

1. Create a new Excel I<workbook> (i.e. file) using C<new()>.

2. Add a I<worksheet> to the new workbook using C<addworksheet()>.

3. Write to the worksheet using C<write()>.

Like this:

    use Spreadsheet::WriteExcel;                            # Step 0

    $workbook  = Spreadsheet::WriteExcel->new("perl.xls");  # Step 1
    $worksheet = $workbook->addworksheet();                 # Step 2
    $worksheet->write('A1', "Hi Excel!");                   # Step 3

This will create an Excel file called C<perl.xls> with a single worksheet and the text C<"Hi Excel!"> in the relevant cell. And that's it. Okay, so there is actually a zeroth step as well, but C<use module> goes without saying. There are also more than 20 examples that come with the distribution and which you can use to get you started. See L<EXAMPLES>.

Those of you who read the instructions first and assemble the furniture afterwards will know how to proceed. ;-)




=head1 WORKBOOK METHODS

The Spreadsheet::WriteExcel module provides an object oriented interface to a new Excel workbook. The following methods are available through a new workbook.

    new()
    close()
    addworksheet($sheetname)
    addformat()
    sheets()
    set_1904()

If you are unfamiliar with object oriented interfaces or the way that they are implemented in Perl have a look at C<perlobj> and C<perltoot> in the main Perl documentation.




=head2 new()

A new Excel workbook is created using the C<new()> constructor which accepts either a filename or a filehandle as a parameter. The following example creates a new Excel file based on a filename:

    my $workbook  = Spreadsheet::WriteExcel->new('filename.xls');
    my $worksheet = $workbook->addworksheet();
    $worksheet->write(0, 0, "Hi Excel!");

Here are some other examples of using C<new()> with filenames:

    my $workbook1 = Spreadsheet::WriteExcel->new($filename);
    my $workbook2 = Spreadsheet::WriteExcel->new("/tmp/filename.xls");
    my $workbook3 = Spreadsheet::WriteExcel->new("c:\\tmp\\filename.xls");
    my $workbook4 = Spreadsheet::WriteExcel->new('c:\tmp\filename.xls');

The last two examples demonstrates how to create a file on DOS or Windows where it is necessary to either escape the directory separator C<\> or to use single quotes to ensure that it isn't interpolated. For more information  see C<perlfaq5: Why can't I use "C:\temp\foo" in DOS paths?>.

The C<new()> constructor returns a Spreadsheet::WriteExcel object that you can use to add worksheets and store data. It should be noted that although C<my> is not specifically required it defines the scope of the new workbook variable and, in the majority of cases, ensures that the workbook is closed properly without explicitly calling the C<close()> method.

If the file cannot be created, due to file permissions or some other reason,  C<new> will return C<undef>. Therefore, it is good practice to check the return value of C<new> before proceeding. As usual the Perl variable C<$!> will be set if there is a file creation error. You will also see one of the warning messages detailed in L<DIAGNOSTICS>:

    my $workbook  = Spreadsheet::WriteExcel->new('protected.xls');
    die "Problems creating new Excel file: $!" unless defined $workbook;

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

Note about the requirement for C<binmode()>: An Excel file is comprised of binary data. Therefore, if you are using a filehandle you should ensure that you C<binmode()> it prior to passing it to C<new()>.You can safely do this regardless of whether your platform requires it or not. For more information about C<binmode()> see C<perlfunc> and C<perlopentut> in the main Perl documentation. It is equally important to note that you do not need to C<binmode()> a filename. In fact it would cause an error. Spreadsheet::WriteExcel performs the C<binmode()> internally when it converts the filename to a filehandle.




=head2 close()

The C<close()> method can be used to explicitly close an Excel file.

    $workbook->close();

An explicit C<close()> is required if the file must be closed prior to performing some external action on it such as copying it, reading its size or attaching it to an email.

In addition, C<close()> may be required if the scope of the Workbook, Worksheet or Format objects cannot be determined by perl. Situations where this can occur are:

=over

=item * If C<my()> was not used to declare the scope of a workbook variable created using C<new()>.

=item * If the C<new()>, C<addworksheet()> or C<addformat()> methods are called in subroutines.

=back

The reason for this is that Spreadsheet::WriteExcel relies on Perl's C<DESTROY> mechanism to trigger destructor methods in a specific sequence. This will not happen if the scope of the variables cannot be determined.


In general, if you create a file with a size of 0 bytes or you fail to create a file you need to call C<close()>.




=head2 addworksheet($sheetname)

At least one worksheet should be added to a new workbook. A worksheet is used to write data into cells:

    $worksheet1 = $workbook->addworksheet();          # Sheet1
    $worksheet2 = $workbook->addworksheet('Foglio2'); # Foglio2
    $worksheet3 = $workbook->addworksheet('Data');    # Data
    $worksheet4 = $workbook->addworksheet();          # Sheet4

If C<$sheetname> is not specified the default Excel convention will be followed, i.e. Sheet1, Sheet2, etc.

Note, you cannot use the same sheet name in more than one worksheet.




=head2 addformat(%properties)

The C<addformat()> method can be used to create new Format objects which are used to apply formatting to a cell. You can either define the properties at creation time via a hash of property values or later via method calls.

    $format1 = $workbook->addformat(%props); # Set properties at creation
    $format2 = $workbook->addformat();       # Set properties later

See the L<CELL FORMATTING> section for more details about Format properties and how to set them.




=head2 sheets()

The C<sheets()> method returns a list of the worksheets in a workbook. This can be useful if you want to repeat an operation on each worksheet in a workbook or if you wish to refer to a worksheet by its index:

    foreach $worksheet ($workbook->sheets()) {
       print $worksheet->get_name();
    }
    
    # or:
    
    ($workbook->sheets())[5]->write('A1', "Hello");


Note: This functionality was previously available via the C<worksheets()> method which returned an array ref. This was unnecessarily complicated. The C<worksheets()> method is still available but deprecated.




=head2 set_1904()

Excel stores dates as real numbers where the integer part stores the number of days since the epoch and the fractional part stores the percentage of the day. The epoch can be either 1900 or 1904. Excel for Windows uses 1900 and Excel for Macintosh uses 1904. However, Excel on either platform will convert automatically between one system and the other.

Spreadsheet::WriteExcel stores dates in the 1900 format by default. If you wish to change this you can call the C<set_1904()> workbook method. You can query the current value by calling the C<get_1904()> workbook method. This returns 0 for 1900 and 1 for 1904.

See also L<Dates in Excel> for more information about working with Excel's date system.

In general you probably won't need to use C<set_1904()>.




=head1 WORKSHEET METHODS

A new worksheet is created by calling the C<addworksheet()> method from a workbook object:

    $worksheet1 = $workbook->addworksheet();
    $worksheet2 = $workbook->addworksheet();

The following methods are available through a new worksheet:

    write()
    write_row()
    write_col()
    write_number()
    write_string()
    write_formula()
    write_blank()
    write_url()
    write_url_range()
    insert_bitmap()
    get_name()
    activate()
    select()
    protect()
    set_first_sheet()
    set_selection()
    set_row()
    set_column()
    freeze_panes()
    thaw_panes()
    merge_cells()
    set_zoom()


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
    $worksheet->write('H2', '=H1+1');

The C<Spreadsheet::WriteExcel::Utility> module that is included in the distro contains helper functions for dealing with A1 notation, for example:

    use Spreadsheet::WriteExcel::Utility;
    
    ($row, $col)    = xl_cell_to_rowcol('C2');  # (1, 2)
    $str            = xl_rowcol_to_cell(1, 2);  # C2

For simplicity, the parameter lists for the worksheet method calls in the following sections are given in terms of row-column notation. In all cases it is also possible to use A1 notation.




=head2 write($row, $column, $token, $format)

Excel distinguishes between data types such as strings, numbers, blanks, formulas and hyperlinks. To simplify the process of writing data C<Spreadsheet::WriteExcel> provides the C<write()> method as a general alias to several more specific methods for writing to a cell in Excel:

    write_string()
    write_number()
    write_blank()
    write_formula()
    write_url()
    write_row()
    write_col()

Here are some examples in both row-column and A1 notation:

                                                      # Same as:
    $worksheet->write(0, 0, "Hello"                ); # write_string()
    $worksheet->write(1, 0, 'One'                  ); # write_string()
    $worksheet->write(2, 0,  2                     ); # write_number()
    $worksheet->write(3, 0,  3.00001               ); # write_number()
    $worksheet->write(4, 0,  ""                    ); # write_blank()
    $worksheet->write(5, 0,  ''                    ); # write_blank()
    $worksheet->write(6, 0,  undef                 ); # write_blank()
    $worksheet->write(7, 0                         ); # write_blank()
    $worksheet->write(8, 0,  'http://www.perl.com/'); # write_url()
    $worksheet->write('A9',  'ftp://ftp.cpan.org/' ); # write_url()
    $worksheet->write('A10', 'internal:Sheet1!A1'  ); # write_url()
    $worksheet->write('A11', 'external:c:\foo.xls' ); # write_url()
    $worksheet->write('A12', '=A3 + 3*A4'          ); # write_formula()
    $worksheet->write('A13', '=SIN(PI()/4)'        ); # write_formula()
    $worksheet->write('A14', \@array               ); # write_row()
    $worksheet->write('A15', [\@array]             ); # write_col()

The general rule is that if it looks like a I<something> then a I<something> is written. The "looks like" is defined by regular expressions:

C<write_number()> if C<$token> is a number based on the following regex: C<$token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/>.

C<write_blank()> if C<$token> is undef or a blank string: C<undef>, C<""> or C<''>.

C<write_url()> if C<$token> is a http, ftp or mailto URL based on the following regexes: C<$token =~ m|^[fh]tt?p://|> or  C<$token =~ m|^mailto:|>.

C<write_url()> if C<$token> is an internal or external sheet reference based on the following regex: C<$token =~ m[^(in|ex)ternal:]>.

C<write_formula()> if the first character of C<$token> is C<"=">.

C<write_row()> if C<$token> is an array ref.

C<write_col()> if C<$token> is an array ref of array refs.

C<write_string()> if none of the previous conditions apply.

The C<$format> parameter is optional. It should be a valid Format object, see L<CELL FORMATTING>:

    my $format = $workbook->addformat();
    $format->set_bold();
    $format->set_color('red');
    $format->set_align('center');

    $worksheet->write(4, 0, "Hello", $format ); # Formatted string

The write() method will ignore empty string or C<undef> tokens unless a format is also supplied. As such you needn't worry about special handling for empty or C<undef> values in your data. See also the the C<write_blank()> method.

One problem with the C<write()> method is that occasionally data looks like a number but you don't want it treated as a number. For example, zip codes or phone numbers often start with a leading zero. If you write it as a number then the leading zero(s) will be stripped. To get around this you can either explicitly write the number as a string or write the number with a number format:

    # write as a number (1209)
    $worksheet->write('A1', '01209');
    
    # Format as a string (01209)
    $worksheet->write_string('A2', '01209');
    
    # Format as a number (01209)
    my $format    = $workbook->addformat(num_format => '00000');
    $worksheet->write('A3', '01209', $format);

Note, Excel writes strings with a left justification and numbers with a right justification so you may want to add an align format as well, L<CELL FORMATTING>.

The C<write> methods return:

    0 for success.
   -1 for insufficient number of arguments.
   -2 for row or column out of bounds.
   -3 for string too long.




=head2 write_number($row, $column, $number, $format)

Write an integer or a float to the cell specified by C<$row> and C<$column>:

    $worksheet->write_number(0, 0,  123456);
    $worksheet->write_number('A2',  2.3451);

See the note about L<Cell notation>. The C<$format> parameter is optional.

In general it is sufficient to use the C<write()> method.




=head2 write_string($row, $column, $string, $format)

Write a string to the cell specified by C<$row> and C<$column>:

    $worksheet->write_string(0, 0, "Your text here" );
    $worksheet->write_string('A2', "or here" );

See the note about L<Cell notation>. The maximum string size is 255 characters. The C<$format> parameter is optional.

In general it is sufficient to use the C<write()> method.




=head2 write_formula($row, $column, $formula, $format)

Write a formula or function to the cell specified by C<$row> and C<$column>:

    $worksheet->write_formula(0, 0, '=$B$3 + B4'  );
    $worksheet->write_formula(1, 0, '=SIN(PI()/4)');
    $worksheet->write_formula(2, 0, '=SUM(B1:B5)' );
    $worksheet->write_formula('A4', '=IF(A3>1,"Yes", "No")'   );
    $worksheet->write_formula('A5', '=AVERAGE(1, 2, 3, 4)'    );
    $worksheet->write_formula('A6', '=DATEVALUE("1-Jan-2001")');

See the note about L<Cell notation>. For more information about writing Excel formulas see L<FORMULAS AND FUNCTIONS IN EXCEL>

In general it is sufficient to use the C<write()> method.




=head2 write_blank($row, $column, $format)

Write a blank cell specified by C<$row> and C<$column>:

    $worksheet->write_blank(0, 0, $format);

This method is used to add formatting to a cell which doesn't contain a string or number value.

Excel differentiates between an "Empty" cell and a "Blank" cell. An "Empty" cell is a cell which doesn't contain data whilst a "Blank" cell is a cell which doesn't contain data but does contain formatting. Excel stores "Blank" cells but ignores "Empty" cells.

As such, if you write an empty cell without formatting it is ignored:

    $worksheet->write('A1',  undef, $format); # write_blank()
    $worksheet->write('A2',  undef         ); # Ignored

This seemingly uninteresting fact means that you can write arrays of data without special treatment for undef or empty string values.

See the note about L<Cell notation>. 




=head2 write_row($row, $column, $array_ref, $format)


The C<write_row()> method can be used to write a 1D or 2D array or data in one go. This is useful for converting the results of a database query into an Excel worksheet. You must pass a reference to the array of data rather than the array itself. The C<write()> method is then called for each element of the data. For example:

    @array      = ('awk', 'gawk', 'mawk');
    $array_ref  = \@array;
    
    $worksheet->write_row(0, 0, $array_ref);
    
    # The above example is equivalent to:
    $worksheet->write(0, 0, $array[0]);
    $worksheet->write(0, 1, $array[1]);
    $worksheet->write(0, 2, $array[2]);


Note: For convenience the C<write()> method behaves in the same way as C<write_row()> if it is passed an array reference. Therefore the following two method calls are equivalent: 

    $worksheet->write_row('A1', $array_ref); # Write a row of data
    $worksheet->write(    'A1', $array_ref); # Same thing

As with all of the write methods the C<$format> parameter is optional. If a format is specified it is applied to all the elements of the data array.

Array references within the data will be treated as columns. This allows you to write 2D arrays of data in one go. For example:

    @eec =  (
                ['maggie', 'milly', 'molly', 'may'  ],
                [13,       14,      15,      16     ],
                ['shell',  'star',  'crab',  'stone']
            );
    
    $worksheet->write_row('A1', \@eec);


Would produce a worksheet as follows:

     -----------------------------------------------------------
    |   |    A    |    B    |    C    |    D    |    E    | ... 
     -----------------------------------------------------------
    | 1 | maggie  | 13      | shell   | ...     |  ...    | ... 
    | 2 | milly   | 14      | star    | ...     |  ...    | ... 
    | 3 | molly   | 15      | crab    | ...     |  ...    | ... 
    | 4 | may     | 16      | stone   | ...     |  ...    | ... 
    | 5 | ...     | ...     | ...     | ...     |  ...    | ... 
    | 6 | ...     | ...     | ...     | ...     |  ...    | ... 


To write the data in a row-column order refer to the C<write_col()> method below.

Any C<undef> values in the data will be ignored unless a format is applied to the data, in which case a formatted blank cell will be written. In either case the appropriate row or column value will still be incremented.

To find out more about array references refer to C<perlref> and C<perlreftut> in the main Perl documentation. To find out more about 2D arrays or "lists of lists" refer to C<perllol>.

The C<write_row()> method returns the first error encountered when writing the elements of the data or zero if no errors were encountered. See the return values described for the C<write()> method above.

See also the C<write_arrays.pl> program in the C<examples> directory of the distro.




=head2 write_col($row, $column, $array_ref, $format)

The C<write_col()> method can be used to write a 1D or 2D array or data in one go. This is useful for converting the results of a database query into an Excel worksheet. You must pass a reference to the array of data rather than the array itself. The C<write()> method is then called for each element of the data. For example:

    @array      = ('awk', 'gawk', 'mawk');
    $array_ref  = \@array;
    
    $worksheet->write_col(0, 0, $array_ref);
    
    # The above example is equivalent to:
    $worksheet->write(0, 0, $array[0]);
    $worksheet->write(1, 0, $array[1]);
    $worksheet->write(2, 0, $array[2]);

As with all of the write methods the C<$format> parameter is optional. If a format is specified it is applied to all the elements of the data array.

Array references within the data will be treated as rows. This allows you to write 2D arrays of data in one go. For example:

    @eec =  (
                ['maggie', 'milly', 'molly', 'may'  ],
                [13,       14,      15,      16     ],
                ['shell',  'star',  'crab',  'stone']
            );
    
    $worksheet->write_col('A1', \@eec);


Would produce a worksheet as follows:

     -----------------------------------------------------------
    |   |    A    |    B    |    C    |    D    |    E    | ... 
     -----------------------------------------------------------
    | 1 | maggie  | milly   | molly   | may     |  ...    | ... 
    | 2 | 13      | 14      | 15      | 16      |  ...    | ... 
    | 3 | shell   | star    | crab    | stone   |  ...    | ... 
    | 4 | ...     | ...     | ...     | ...     |  ...    | ... 
    | 5 | ...     | ...     | ...     | ...     |  ...    | ... 
    | 6 | ...     | ...     | ...     | ...     |  ...    | ... 


To write the data in a column-row order refer to the C<write_row()> method above.

Any C<undef> values in the data will be ignored unless a format is applied to the data, in which case a formatted blank cell will be written. In either case the appropriate row or column value will still be incremented.

As noted above the C<write()> method can be used as a synonym for C<write_row()> and C<write_row()> handles nested array refs as columns. Therefore, the following two method calls are equivalent although the more explicit call to C<write_col()> would be preferable for maintainability: 

    $worksheet->write_col('A1', $array_ref    ); # Write a column of data
    $worksheet->write(    'A1', [ $array_ref ]); # Same thing

To find out more about array references refer to C<perlref> and C<perlreftut> in the main Perl documentation. To find out more about 2D arrays or "lists of lists" refer to C<perllol>.

The C<write_col()> method returns the first error encountered when writing the elements of the data or zero if no errors were encountered. See the return values described for the C<write()> method above.

See also the C<write_arrays.pl> program in the C<examples> directory of the distro.




=head2 write_url($row, $col, $url, $string, $format)

Write a hyperlink to a URL in the cell specified by C<$row> and C<$column>. The hyperlink is comprised of two elements: the visible label and the invisible link. The visible label is the same as the link unless an alternative string is specified. The parameters C<$string> and the C<$format> are optional and their position is interchangeable.

The label is written using the C<write_string()> method. Therefore the 255 characters string limit applies to the label: the URL can be any length.

There are three web style URI's supported: C<http://>, C<ftp://> and  C<mailto:>:

    $worksheet->write_url(0, 0,  'ftp://www.perl.org/'                  );
    $worksheet->write_url(1, 0,  'http://www.perl.com/', 'Perl home'    );
    $worksheet->write_url('A3',  'http://www.perl.com/', $format        );
    $worksheet->write_url('A4',  'http://www.perl.com/', 'Perl', $format);
    $worksheet->write_url('A5',  'mailto:jmcnamara@cpan.org'            );

There are two local URIs supported: C<internal:> and C<external:>. These are used for hyperlinks to internal worksheet references or external workbook and worksheet references:

    $worksheet->write_url('A6',  'internal:Sheet2!A1'                   );
    $worksheet->write_url('A7',  'internal:Sheet2!A1',   $format        );
    $worksheet->write_url('A8',  'internal:Sheet2!A1:B2'                );
    $worksheet->write_url('A9',  q{internal:'Sales Data'!A1}            );
    $worksheet->write_url('A10', 'external:c:\temp\foo.xls'             );
    $worksheet->write_url('A11', 'external:c:\temp\foo.xls#Sheet2!A1'   );
    $worksheet->write_url('A12', 'external:..\..\..\foo.xls'            );
    $worksheet->write_url('A13', 'external:..\..\..\foo.xls#Sheet2!A1'  );
    $worksheet->write_url('A13', 'external:\\\\NETWORK\share\foo.xls'   );

All of the these URI types are recognised by the C<write()> method, see above.

Worksheet references are typically of the form C<Sheet1!A1>. You can also refer to a worksheet range using the standard Excel notation: C<Sheet1!A1:B2>.

In external links the workbook and worksheet name must be separated by the C<#> character: C<external:Workbook.xls#Sheet1!A1'>.

You can also link to a named range in the target worksheet. For example say you have a named range called C<my_name> in the workbook C<c:\temp\foo.xls> you could link to it as follows:

    $worksheet->write_url('A14', 'external:c:\temp\foo.xls#my_name');

Note, you cannot currently create named ranges with C<Spreadsheet::WriteExcel>. 

Excel requires that worksheet names containing spaces or non alphanumeric characters are single quoted as follows C<'Sales Data'!A1>. If you need to do this in a single quoted string then you can either escape the single quotes C<\'> or use the quote operator C<q{}> as described in C<perlop> in the main Perl documentation.

Links to network files are also supported. MS/Novell Network files normally begin with two back slashes as follows C<\\NETWORK\etc>. In order to generate this in a single or double quoted string you will have to escape the backslashes,  C<'\\\\NETWORK\etc'>.

If you are using double quote strings then you should be careful to escape anything that looks like a metacharacter. For more information  see C<perlfaq5: Why can't I use "C:\temp\foo" in DOS paths?>.

Finally, you can avoid most of these quoting problems by using forward slashes. These are translated internally to backslashes:

    $worksheet->write_url('A14', "external:c:/temp/foo.xls"             );
    $worksheet->write_url('A15', 'external://NETWORK/share/foo.xls'     );

Note: Hyperlinks are not available in Excel 5. They will appear as a string only.

See also, the note about L<Cell notation>.




=head2 write_url_range($row1, $col1, $row2, $col2, $url, $string, $format)

This method is essentially the same as the C<write_url()> method described above. The main difference is that you can specify that the link is available for a range of cells:

    $worksheet->write_url(0, 0, 0, 3, 'ftp://www.perl.org/'              );
    $worksheet->write_url(1, 0, 0, 3, 'http://www.perl.com/', 'Perl home');
    $worksheet->write_url('A3:D3',    'internal:Sheet2!A1'               );
    $worksheet->write_url('A4:D4',     'external:c:\temp\foo.xls'        );


This method is generally only required when used in conjunction with merged cells. See the C<merge_cells()> method and the C<merge> property of a Format object, L<CELL FORMATTING>.

There is no way to force this behaviour through the C<write()> method.

The parameters C<$string> and the C<$format> are optional and their position is interchangeable. However, they are applied only to the first cell in the range. 

Note: Hyperlinks are not available in Excel 5. They will appear as a string only.

See also, the note about L<Cell notation>.




=head2 insert_bitmap($row, $col, $filename, $x, $y, $scale_x, $scale_y)

This method can be used to insert a bitmap into a worksheet. The bitmap must be a 24 bit, true colour, bitmap. No other format is supported. The C<$x>, C<$y>, C<$scale_x> and C<$scale_y> parameters are optional.

    $worksheet1->insert_bitmap('A1', 'perl.bmp');
    $worksheet2->insert_bitmap('A1', '../images/perl.bmp');
    $worksheet3->insert_bitmap('A1', '.c:\images\perl.bmp');

Note: you must call C<set_row()> or C<set_column()> before C<insert_bitmap()> if you wish to change the default dimensions of any of the rows or columns that the images occupies. Also, if you use large fonts then the height of the row that they occupy may change automatically. This in turn could affect the scaling of your image. To avoid this you should explicitly set the height of the row using C<set_row()>.

The parameters C<$x> and C<$y> can be used to specify an offset from the top left hand corner of the the cell specified by C<$row> and C<$col>. The offset values are in pixels.

    $worksheet1->insert_bitmap('A1', 'perl.bmp', 32, 10);

The default width of a cell is 63 pixels. The default height of a cell is 17 pixels. The offsets are ignored if they are greater than the width or height of the underlying cell.

The pixels offsets can be calculated using the following relationships:

    Wp = 7We +5 
    Hp = 4/3He

    where:
    We is the cell width in Excels units
    Wp is width in pixels
    He is the cell height in Excels units
    Hp is height in pixels

The parameters C<$scale_x> and C<$scale_y> can be used to scale the inserted image horizontally and vertically:

    # Scale the inserted image: width x 2.0, height x 0.8
    $worksheet->insert_bitmap('A1', 'perl.bmp', 0, 0, 2, 0.8); 

Note: although Excel allows you to import several graphics formats such as gif, jpeg, png and eps these are converted internally into a proprietary format. One of the few non-proprietary formats that Excel supports is 24 bit, true colour, bitmaps. Therefore if you wish to use images in any other format you must first use an external application such as ImageMagick/Perl::Magick to convert them to 24 bit bitmaps.

A later release will support the use of file handles and pre-encoded bitmap strings.

See also the C<images.pl> program in the C<examples> directory of the distro.




=head2 get_name()

The C<get_name()> method is used to retrieve the name of a worksheet. For example:

    foreach my $sheet ($workbook->sheets()) {
        print $sheet->get_name();
    }




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

A selected worksheet has its tab highlighted. Selecting worksheets is a way of grouping them together so that, for example, several worksheets could be printed in one go. A worksheet that has been activated via the C<activate()> method will also appear as selected. You probably won't need to use the C<select()> method very often.




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




=head2 protect($password)

The C<protect()> method is used to protect a worksheet from modification:

    $worksheet->protect();

It can be turned off in Excel via the C<Tools-E<gt>Protection-E<gt>Unprotect Sheet> menu command.

The C<protect()> method also has the effect of enabling a cell's C<locked> and C<hidden> properties if they have been set. A "locked" cell cannot be edited. A "hidden" cell will display the results of a formula but not the formula itself. In Excel a cell's locked property is on by default.

    # Set some format properties
    my $unlocked  = $workbook->addformat(locked => 0);
    my $hidden    = $workbook->addformat(hidden => 1);
    
    # Enable worksheet protection
    $worksheet->protect();
    
    # This cell cannot be edited, it is locked by default
    $worksheet->write('A1', '=1+2');
    
    # This cell can be edited
    $worksheet->write('A2', '=1+2', $unlocked);
    
    # The formula in this cell isn't visible 
    $worksheet->write('A3', '=1+2', $hidden);

See also the C<set_locked> and C<set_hidden> format methods in L<CELL FORMATTING>.

You can optionally add a password to the worksheet protection:

    $worksheet->protect('drowssap');

Note, the worksheet level password in Excel provides very weak protection. It does not encrypt your data in any way and it is very easy to deactivate. Therefore, do not use the above method if you wish to protect sensitive data or calculations. However, before you get worried, Excel's own workbook level password protection does provide strong encryption in Excel 97+. For reasons both ethical and technical, this will never be supported by C<Spreadsheet::WriteExcel>.




=head2 set_selection($first_row, $first_col, $last_row, $last_col)

This method can be used to specify which cell or cells are selected in a worksheet. The most common requirement is to select a single cell, in which case C<$last_row> and C<$last_col> can be omitted. The active cell within a selected range is determined by the order in which C<$first> and C<$last> are specified. It is also possible to specify a cell or a range using A1 notation. See the note about L<Cell notation>.

Examples:

    $worksheet1->set_selection(3, 3);       # 1. Cell D4.
    $worksheet2->set_selection(3, 3, 6, 6); # 2. Cells D4 to G7.
    $worksheet3->set_selection(6, 6, 3, 3); # 3. Cells G7 to D4.
    $worksheet4->set_selection('D4');       # Same as 1.
    $worksheet5->set_selection('D4:G7');    # Same as 2.
    $worksheet6->set_selection('G7:D4');    # Same as 3.

The default cell selections is (0, 0), 'A1'.




=head2 set_row($row, $height, $format)

This method can be used to specify the height of a row. The C<$format> parameter is optional, for additional information, see L<CELL FORMATTING>.

    $worksheet->set_row(0, 20); # Row 1 height set to 20

If you wish to set the format without changing the height you can pass C<undef> as the height parameter:

    $worksheet->set_row(0, undef, $format);

The C<$format> parameter will only define a format if C<set_row()> is called after the cells have been written:

    $worksheet->write('A1', "Hello");       # Formatted
    $worksheet->set_row(0, undef, $format);
    $worksheet->write('B1', "Hello");       # Not formatted

This behaviour will be fixed in a future release.




=head2 set_column($first_col, $last_col, $width, $format, $hidden)

This method can be used to specify the width of a single column or a range of columns. If the method is applied to a single column the value of C<$first_col> and C<$last_col> should be the same. It is also possible to specify a column range using the form of A1 notation used for columns. See the note about L<Cell notation>.

Examples:

    $worksheet->set_column(0, 0,  20); # Column  A   width set to 20
    $worksheet->set_column(1, 3,  30); # Columns B-D width set to 30
    $worksheet->set_column('E:E', 20); # Column  E   width set to 20
    $worksheet->set_column('F:H', 30); # Columns F-H width set to 30

The width corresponds to the column width value that is specified in Excel. It is approximately equal to the length of a string in the default font of Arial 10. Unfortunately, there is no way to specify "AutoFit" for a column in the Excel file format. This feature is only available at runtime from within Excel.

The C<$format> parameter is optional, for additional information, see L<CELL FORMATTING>. If you wish to set the format without changing the width you can pass C<undef> as the width parameter:

    $worksheet->set_column(0, 0, undef, $format);

The C<$format> parameter will not set the format for individual cells written by Spreadsheet::WriteExcel, it only has an effect on cells written after the workbook is opened in Excel. This behaviour will be fixed in a future release.

The C<$hidden> parameter is optional. It should be set to 1 if you wish to hide a column. This can be used, for example, to hide intermediary steps in a complicated calculation:

    $worksheet->set_column('D:D', 20,    $format, 1);
    $worksheet->set_column('E:E', undef, undef,   1);




=head2 freeze_panes($row, $col, $top_row, $left_col)

This method can be used to divide a worksheet into horizontal or vertical regions known as panes and to also "freeze" these panes so that the splitter bars are not visible. This is the same as the C<Window-E<gt>Freeze Panes> menu command in Excel

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
    

See also the C<panes.pl> program in the C<examples> directory of the distribution.




=head2 thaw_panes($y, $x, $top_row, $left_col)

This method can be used to divide a worksheet into horizontal or vertical regions known as panes. This method is different from the C<freeze_panes()> method in that the splits between the panes will be visible to the user and each pane will have its own scroll bars.

The parameters C<$y> and C<$x> are used to specify the vertical and horizontal position of the split. The units for C<$y> and C<$x> are the same as those used by Excel to specify row height and column width. However, the vertical and horizontal units are different from each other. Therefore you must specify the C<$y> and C<$x> parameters in terms of the row heights and column widths that you have set or the default values which are C<12.75> for a row and  C<8.43> for a column.

You can set one of the C<$y> and C<$x> parameters as zero if you do not want either a vertical or horizontal split. The parameters C<$top_row> and C<$left_col> are optional. They are used to specify the top-most or left-most visible row or column in the bottom-right pane.

Example:

    $worksheet->thaw_panes(12.75, 0,    1, 0); # First row
    $worksheet->thaw_panes(0,     8.43, 0, 1); # First column
    $worksheet->thaw_panes(12.75, 8.43, 1, 1); # First row and column

You cannot use A1 notation with this method.

See also the C<freeze_panes()> method and the C<panes.pl> program in the C<examples> directory of the distribution.




=head2 merge_cells($first_row, $first_col, $last_row, $last_col)

Merging cells is generally achieved by setting the C<merge> property of a Format object, see L<CELL FORMATTING>. However, in certain circumstances this is not sufficient and you must additionally specify the cells to be merged via the C<merge_cells()> method.

The main use of the C<merge_cells()> method is to merge cells vertically.
    
The C<merge_cells()> method can also be used to merge cells that contain hyperlinks although this can also be achieved via the C<write_url_range()> method.

For an example of how to use this method see the C<merge3.pl> program in the C<examples> directory of the distribution.

This method is currently of limited use. It will play a more important role when Spreadsheet::WriteExcel moves to the Excel 97/2000 file format.

In general the C<set_merge()> method is all that you will require to create merged cells, see L<CELL FORMATTING>.



=head2 set_zoom($scale)

Set the worksheet zoom factor in the range C<10 E<lt>= $scale E<lt>= 400>:

    $worksheet1->set_zoom(50);
    $worksheet2->set_zoom(75);
    $worksheet3->set_zoom(300);
    $worksheet4->set_zoom(400);

The default zoom factor is 100. You cannot zoom to "Selection" because it is calculated by Excel at run-time.

Note, C<set_zoom()> does not affect the scale of the printed page. For that you should use C<set_print_scale()>.




=head1 PAGE SET-UP METHODS

Page set-up methods affect the way that a worksheet looks when it is printed. They control features such as page headers and footers and margins. These methods are really just standard worksheet methods. They are documented here in a separate section for the sake of clarity.

The following methods are available for page set-up:

    set_landscape()
    set_portrait()
    set_paper()
    center_horizontally()
    center_vertically()
    set_margins()
    set_header()
    set_footer()
    repeat_rows()
    repeat_columns()
    hide_gridlines()
    print_row_col_headers()
    print_area()
    fit_to_pages()
    set_print_scale()
    set_h_pagebreaks()
    set_v_pagebreaks()


A common requirement when working with Spreadsheet::WriteExcel is to apply the same page set-up features to all of the worksheets in a workbook. To do this you can use the C<sheets()> method of the C<workbook> class to access the array of worksheets in a workbook:

    foreach $worksheet ($workbook->sheets()) {
       $worksheet->set_landscape();
    }




=head2 set_landscape()

This method is used to set the orientation of a worksheet's printed page to landscape:

    $worksheet->set_landscape(); # Landscape mode




=head2 set_portrait()

This method is used to set the orientation of a worksheet's printed page to portrait. The default worksheet orientation is portrait, so you won't generally need to call this method.

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

    $worksheet->set_paper(1); # US Letter
    $worksheet->set_paper(9); # A4

If you do not specify a paper type the worksheet will print using the printer's default paper. 




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

    Control             Category            Description
    =======             ========            ===========
    &L                  Justification       Left
    &C                                      Center
    &R                                      Right
    
    &P                  Information         Page number
    &N                                      Total number of pages
    &D                                      Date
    &T                                      Time
    &F                                      File name
    &A                                      Worksheet name
    
    &fontsize           Font                Font size
    &"font,style"                           Font name and style
    &U                                      Single underline
    &E                                      Double underline
    &S                                      Strikethrough
    &X                                      Superscript
    &Y                                      Subscript
    
    &&                  Miscellaneous       Literal ampersand &


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
    

If you do not specify any justification the text will be centred:

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

    $worksheet1->set_header('&30Hello Big'  );
    $worksheet2->set_header('&10Hello Small');

You can specify the font of a section of the text by prefixing it with the control sequence C<&"font,style"> where C<fontname> is a font name such as "Courier New" or "Times New Roman" and C<style> is one of the standard Windows font descriptions: "Regular", "Italic", "Bold" or "Bold Italic":

    $worksheet1->set_header('&"Courier New,Italic"Hello');
    $worksheet2->set_header('&"Courier New,Bold Italic"Hello');
    $worksheet3->set_header('&"Times New Roman,Regular"Hello');

It is possible to combine all of these features together to create sophisticated headers and footers. As an aid to setting up complicated headers and footers you can record a page set-up as a macro in Excel and look at the format strings that VBA produces. Remember however that VBA uses two double quotes C<""> to indicate a single double quote. For the last example above the equivalent VBA code looks like this:

    .LeftHeader   = ""
    .CenterHeader = "&""Times New Roman,Regular""Hello"
    .RightHeader  = ""


To include a single literal ampersand C<&> in a header or footer you should use a double ampersand C<&&>:

    $worksheet1->set_header('Rhythm && Blues');

As stated above the margin parameter is optional. As with the other margins the value should be in inches. The default header and footer margin is 0.50 inch. The header and footer margin size can be set as follows:

    $worksheet->set_header('&CHello', 0.75);

The header and footer margins are independent of the top and bottom margins.

Note, the header or footer string must be less than 255 characters. Strings longer than this will not be written and a warning will be generated.




=head2 set_footer()

Use of the C<set_footer()> method is the same as the C<set_header()> method explained above.




=head2 repeat_rows($first_row, $last_row)

Set the number of rows to repeat at the top of each printed page.

For large Excel documents it is often desirable to have the first row or rows of the worksheet print out at the top of each page. This can be achieved by using the C<repeat_rows()> method. The parameters C<$first_row> and C<$last_row> are zero based. The C<$last_row> parameter is optional if you only wish to specify one row:

    $worksheet1->repeat_rows(0);    # Repeat the first row
    $worksheet2->repeat_rows(0, 1); # Repeat the first two rows




=head2 repeat_columns($first_col, $last_col)

Set the columns to repeat at the left hand side of each printed page.

For large Excel documents it is often desirable to have the first column or columns of the worksheet print out at the left hand side of each page. This can be achieved by using the C<repeat_columns()> method. The parameters C<$first_column> and C<$last_column> are zero based. The C<$last_column> parameter is optional if you only wish to specify one column. You can also specify the columns using A1 column notation, see the note about L<Cell notation>.

    $worksheet1->repeat_columns(0);     # Repeat the first column
    $worksheet2->repeat_columns(0, 1);  # Repeat the first two columns
    $worksheet3->repeat_columns('A:A'); # Repeat the first column
    $worksheet4->repeat_columns('A:B'); # Repeat the first two columns




=head2 hide_gridlines()

This method is used to hide the gridlines on a printed page. 

Gridlines are the lines that divide the cells on a worksheet. Printed gridlines are turned on by default. If you have defined your own cell borders you may wish to hide the gridlines on the printed page.

    $worksheet->hide_gridlines();




=head2 print_row_col_headers()

Set the option to print the row and column headers on the printed page.

An Excel worksheet looks something like the following;

     ------------------------------------------
    |   |   A   |   B   |   C   |   D   |  ... 
     ------------------------------------------
    | 1 |       |       |       |       |  ... 
    | 2 |       |       |       |       |  ... 
    | 3 |       |       |       |       |  ... 
    | 4 |       |       |       |       |  ... 
    |...|  ...  |  ...  |  ...  |  ...  |  ... 

The headers are the letters and numbers at the top and the left of the worksheet. Since these headers serve mainly as a indication of position on the worksheet they generally do not appear on the printed page. If you wish to have them printed you can use the C<print_row_col_headers()> method : 

    $worksheet->print_row_col_headers()

Do not confuse these headers with page headers as described in the C<set_header()> section above.




=head2 print_area($first_row, $first_col, $last_row, $last_col)

This method is used to specify the area of the worksheet that will be printed. All four parameters must be specified. You can also use A1 notation, see the note about L<Cell notation>.


    $worksheet1->print_area("A1:H20");    # Cells A1 to H20
    $worksheet2->print_area(0, 0, 19, 7); # The same




=head2 fit_to_pages($width, $height)

The C<fit_to_pages()> method is used to fit the printed area to a specific number of pages both vertically and horizontally. If the printed area exceeds the specified number of pages it will be scaled down to fit. This guarantees that the printed area will always appear on the specified number of pages even if the page size or margins change.

    $worksheet1->fit_to_pages(1, 1); # Fit to 1x1 pages
    $worksheet2->fit_to_pages(2, 1); # Fit to 2x1 pages
    $worksheet3->fit_to_pages(1, 2); # Fit to 1x2 pages

The print area can be defined using the C<print_area()> method as described above. 

A common requirement is to fit the printed output to "n" pages wide but have the height be as long as necessary. To achieve this set the C<$height> to zero or leave it blank:

    $worksheet1->fit_to_pages(1, 0); # 1 page wide and as long as necessary
    $worksheet2->fit_to_pages(1);    # The same
    

Note that although it is valid to use both C<fit_to_pages()> and C<set_print_scale()> on the same worksheet only one of these options can be active at a time. The last method call made will set the active option.

Note that C<fit_to_pages()> will override any manual page breaks that are defined in the worksheet.




=head2 set_print_scale($scale)

Set the scale factor of the printed page. Scale factors in the range C<10 E<lt>= $scale E<lt>= 400> are valid:

    $worksheet1->set_print_scale(50);
    $worksheet2->set_print_scale(75);
    $worksheet3->set_print_scale(300);
    $worksheet4->set_print_scale(400);

The default scale factor is 100. Note, C<set_print_scale()> does not affect the scale of the visible page in Excel. For that you should use C<set_zoom()>.

Note also that although it is valid to use both C<fit_to_pages()> and C<set_print_scale()> on the same worksheet only one of these options can be active at a time. The last method call made will set the active option.




=head2 set_h_pagebreaks(@breaks)

Add horizontal page breaks to a worksheet. A page break causes all the data that follows it to be printed on the next page. Horizontal page breaks act between rows. To create a page break between rows 20 and 21 you must specify the break at row 21. However in zero index notation this is actually row 20. So you can pretend for a small while that you are using 1 index notation:

    $worksheet1->set_h_pagebreaks(20); # Break between row 20 and 21

The C<set_h_pagebreaks()> method will accept a list of page breaks and you can call it more than once:

    $worksheet2->set_h_pagebreaks( 20,  40,  60,  80, 100); # Add breaks
    $worksheet2->set_h_pagebreaks(120, 140, 160, 180, 200); # Add some more

Note: If you specify the "fit to page" option via the C<fit_to_pages()> method it will override all manual page breaks.

There is a silent limitation of about 1000 horizontal page breaks per worksheet in line with an Excel internal limitation.




=head2 set_v_pagebreaks(@breaks)

Add vertical page breaks to a worksheet. A page break causes all the data that follows it to be printed on the next page. Vertical page breaks act between columns. To create a page break between columns 20 and 21 you must specify the break at column 21. However in zero index notation this is actually column 20. So you can pretend for a small while that you are using 1 index notation:

    $worksheet1->set_v_pagebreaks(20); # Break between column 20 and 21

The C<set_v_pagebreaks()> method will accept a list of page breaks and you can call it more than once:

    $worksheet2->set_v_pagebreaks( 20,  40,  60,  80, 100); # Add breaks
    $worksheet2->set_v_pagebreaks(120, 140, 160, 180, 200); # Add some more

Note: If you specify the "fit to page" option via the C<fit_to_pages()> method it will override all manual page breaks.




=head1 CELL FORMATTING

This section describes the methods and properties that are available for formatting cells in Excel. The properties of a cell that can be formatted include: fonts, colours, patterns, borders, alignment and number formatting.


=head2 Creating and using a Format object

Cell formatting is defined through a Format object. Format objects are created by calling the workbook C<addformat()> method as follows:

    my $format1 = $workbook->addformat();       # Set properties later
    my $format2 = $workbook->addformat(%props); # Set properties at creation

The format object holds all the formatting properties that can be applied to a cell, a row or a column. The process of setting these properties is discussed in the next section.

Once a Format object has been constructed and it properties have been set it can be passed as an argument to the worksheet C<write> methods as follows:

    $worksheet->write(0, 0, "One", $format);
    $worksheet->write_string(1, 0, "Two", $format);
    $worksheet->write_number(2, 0, 3, $format);
    $worksheet->write_blank(3, 0, $format);

Formats can also be passed to the worksheet C<set_row()> and C<set_column()> methods to define the default property for a row or column.

    $worksheet->set_row(0, 15, $format);
    $worksheet->set_column(0, 0, 15, $format);

However, the C<set_row()> and C<set_column()> methods will not set the format for individual cells written by WriteExcel, they only have an effect on cells written after the workbook is opened in Excel.

=head2 Format methods and Format properties

The following table shows the Excel format categories, the formatting properties that can be applied and the equivalent object method:


    Category   Description       Property        Method Name         
    --------   -----------       --------        -----------         
    Font       Font type         font            set_font()          
               Font size         size            set_size()          
               Font color        color           set_color()         
               Bold              bold            set_bold()          
               Italic            italic          set_italic()        
               Underline         underline       set_underline()     
               Strikeout         font_strikeout  set_font_strikeout()
               Super/Subscript   font_script     set_font_script()   
               Outline           font_outline    set_font_outline()  
               Shadow            font_shadow     set_font_shadow()   
                                                                     
    Number     Numeric format    num_format      set_num_format()    
                                                                     
    Protection Lock cells        locked          set_locked()    
               Hide formulas     hidden          set_hidden()    
                                                                     
    Alignment  Horizontal align  align           set_align()         
               Vertical align    valign          set_align()         
               Rotation          rotation        set_rotation()      
               Text wrap         text_wrap       set_text_wrap()     
               Justify last      text_justlast   set_text_justlast() 
               Merge             merge           set_merge()         
                                                                     
    Pattern    Cell pattern      pattern         set_pattern()       
               Background color  bg_color        set_bg_color()      
               Foreground color  fg_color        set_fg_color()      
                                                                     
    Border     Cell border       border          set_border()        
               Bottom border     bottom          set_bottom()        
               Top border        top             set_top()           
               Left border       left            set_left()          
               Right border      right           set_right()         
               Border color      border_color    set_border_color()  
               Bottom color      bottom_color    set_bottom_color()  
               Top color         top_color       set_top_color()     
               Left color        left_color      set_left_color()    
               Right color       right_color     set_right_color()   

There are two ways of setting Format properties: by using the object method interface or by setting the property directly. For example, a typical use of the method interface would be as follows:

    my $format = $workbook->addformat();
    $format->set_bold();
    $format->set_color('red');

By comparison the properties can be set directly by passing a hash of properties to the Format constructor:

    my $format = $workbook->addformat(bold => 1, color => 'red');

or after the Format has been constructed by means of the C<set_properties()> method as follows:

    my $format = $workbook->addformat();
    $format->set_properties(bold => 1, color => 'red');

You can also store the properties in one or more named hashes and pass them to the required method:

    my %font    = (
                    font  => 'Arial',
                    size  => 12,
                    color => 'blue',
                    bold  => 1,
                  );

    my %shading = (
                    fg_color => 'green',
                    pattern  => 1,
                  );


    my $format1 = $workbook->addformat(%font);            # Font only
    my $format2 = $workbook->addformat(%font, %shading);  # Font and shading


The provision of two ways of setting properties might lead you to wonder which is the best way. The answer depends on the amount of formatting that will be required in your program. Initially, Spreadsheet::WriteExcel only allowed individual Format properties to be set via the appropriate method. While this was sufficient for most circumstances it proved very cumbersome in programs that required a large amount of formatting. In addition the mechanism for reusing properties between Format objects was complicated.

As a result the Perl/Tk style of adding properties was added to, hopefully, facilitate developers who need to define a lot of formatting. In fact the Tk style of defining properties is also supported:

    my %font    = (
                    -font      => 'Arial',
                    -size      => 12,
                    -color     => 'blue',
                    -bold      => 1,
                  );

An additional advantage of working with hashes of properties is that it allows you to share formatting between workbook objects

You can also create a format "on the fly" and pass it directly to a write method as follows:

    $worksheet->write('A1', "Title", $workbook->addformat(bold => 1));

This corresponds to an "anonymous" format in the Perl sense of anonymous data or subs.

If you need to create an Excel file with a large amount of formatting you can also use the C<lecxe.pl> program in the C<examples> directory of the distribution. C<lecxe> is a Win32::OLE program written by Tomas Andersson which converts Excel files to Spreadsheet::WriteExcel files. Therefore, you can use Excel to define your formatting and have C<lecxe> do the hard work for you.




=head2 Working with formats

The default format is Arial 10 with all other properties off. 

Each unique format in Spreadsheet::WriteExcel must have a corresponding Format object. It isn't possible to use a Format with a write() method and then redefine the Format for use at a later stage. This is because a Format is applied to a cell not in its current state but in its final state. Consider the following example:

    my $format = $workbook->addformat();
    $format->set_bold();
    $format->set_color('red');
    $worksheet->write('A1', "Cell A1", $format);
    $format->set_color('green');
    $worksheet->write('B1', "Cell B1", $format);

Cell A1 is assigned the Format C<$format> which is initially set to the colour red. However, the colour is subsequently set to green. When Excel displays Cell A1 it will display the final state of the Format which in this case will be the colour green.

In general a method call without an argument will turn a property on, for example:

    my $format1 = $workbook->addformat();
    $format1->set_bold();  # Turns bold on
    $format1->set_bold(1); # Also turns bold on
    $format1->set_bold(0); # Turns bold off




=head1 FORMAT METHODS

The Format object methods are described in more detail in the following sections. In addition, there is a Perl program called C<formats.pl> in the C<examples> directory of the WriteExcel distribution. This program creates an Excel workbook called C<formats.xls> which contains examples of almost all the format types.

The following Format methods are available:

    set_font()          
    set_size()          
    set_color()         
    set_bold()          
    set_italic()        
    set_underline()     
    set_font_strikeout()
    set_font_script()   
    set_font_outline()  
    set_font_shadow()   
    set_num_format()
    set_locked()
    set_hidden()
    set_align()         
    set_align()         
    set_rotation()      
    set_text_wrap()     
    set_text_justlast() 
    set_merge()         
    set_pattern()       
    set_bg_color()      
    set_fg_color()      
    set_border()        
    set_bottom()        
    set_top()           
    set_left()          
    set_right()         
    set_border_color()  
    set_bottom_color()  
    set_top_color()     
    set_left_color()    
    set_right_color()   


The above methods can also be applied directly as properties. For example C<$worksheet-E<gt>set_bold()> is equivalent to C<set_properties(bold =E<gt> 1)>.


=head2 set_properties(%properties)

The properties of an existing Format object can be set by means of C<set_properties()>:

    my $format = $workbook->addformat();
    $format->set_properties(bold => 1, color => 'red');

You can also store the properties in one or more named hashes and pass them to the C<set_properties()> method:

    my %font    = (
                    font  => 'Arial',
                    size  => 12,
                    color => 'blue',
                    bold  => 1,
                  );

    my $format = $workbook->set_properties(%font);

This method can be used as an alternative to setting the properties with C<addformat()> or the specific format methods that are detailed in the following sections.




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
    Valid args:         Integers from 8..63 or the following strings:
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

For additional examples see the 'Named colors' and 'Standard colors' worksheets created by formats.pl in the examples directory.




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

Set the italic property of the font:

    $format->set_italic();  # Turn italic on




=head2 set_underline()

    Default state:      Underline is off
    Default action:     Turn on single underline
    Valid args:         0  = No underline
                        1  = Single underline
                        2  = Double underline
                        33 = Single accounting underline
                        34 = Double accounting underline

Set the underline property of the font.

    $format->set_underline();   # Single underline




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

    # Zip code
    $format13->set_num_format('00000');
    $worksheet->write(14, 0, '01209',   $format13);


The number system used for dates is described in L<Dates in Excel>. 

The colour format should have one of the following values:

    [Black] [Blue] [Cyan] [Green] [Magenta] [Red] [White] [Yellow]

Alternatively you can specify the colour based on a colour index as follows: C<[Color n]>, where n is a standard Excel colour index - 7. See the 'Standard colors' worksheet created by formats.pl.

For more information refer to the documentation on formatting in the C<doc> directory of the Spreadsheet::WriteExcel distro, the Excel on-line help or to the tutorial at: http://support.microsoft.com/support/Excel/Content/Formats/default.asp and http://support.microsoft.com/support/Excel/Content/Formats/codes.asp

You should ensure that the format string is valid in Excel prior to using it in WriteExcel.

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




=head2 set_locked()

    Default state:      Cell locking is on
    Default action:     Turn locking on
    Valid args:         0, 1

This property can be used to prevent modification of a cells contents. Following Excel's convention, cell locking is turned on by default. However, it only has an effect if the worksheet has been protected, see the worksheet C<protect()> method.

    my $locked  = $workbook->addformat();
    $locked->set_locked(1); # A non-op
    
    my $unlocked = $workbook->addformat();
    $locked->set_locked(0);
    
    # Enable worksheet protection
    $worksheet->protect();
    
    # This cell cannot be edited.
    $worksheet->write('A1', '=1+2', $locked);
    
    # This cell can be edited.
    $worksheet->write('A2', '=1+2', $unlocked);

Note: This offers weak protection even with a password, see the note in relation to the C<protect()> method.




=head2 set_hidden()

    Default state:      Formula hiding is off
    Default action:     Turn hiding on
    Valid args:         0, 1

This property is used to hide a formula while still displaying its result. This is generally used to hide complex calculations from end users who are only interested in the result. It only has an effect if the worksheet has been protected, see the worksheet C<protect()> method.

    my $hidden = $workbook->addformat();
    $hidden->set_hidden();
    
    # Enable worksheet protection
    $worksheet->protect();
    
    # The formula in this cell isn't visible 
    $worksheet->write('A1', '=1+2', $hidden);


Note: This offers weak protection even with a password, see the note in relation to the C<protect()> method.




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

See also the C<merge1.pl>, C<merge2.pl> and C<merge3.pl> programs in the C<examples> directory and the C<merge_cells()> method.



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


Set the rotation of the text in a cell. See the 'Alignment' worksheet created by formats.pl. Note, fractional rotations aren't possible with the Excel 5 format.





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




=head2 copy($format)


This method is used to copy all of the properties from one Format object to another:

    my $lorry1 = $workbook->addformat();
    $lorry1->set_bold();
    $lorry1->set_italic();
    $lorry1->set_color('red');    # lorry1 is bold, italic and red

    my $lorry2 = $workbook->addformat();
    $lorry2->copy($lorry1);
    $lorry2->set_color('yellow'); # lorry2 is bold, italic and yellow

It is only useful if you are using the method interface to Format properties. It generally isn't required if you are setting Format properties directly using hashes.


Note: this is not a copy constructor, both objects must exist prior to copying.




=head1 DATES IN EXCEL


Dates and times in Excel are represented by real numbers, for example "Jan 1 2001 12:30 AM" is represented by the number 36892.521.

The integer part of the number stores the number of days since the epoch and the fractional part stores the percentage of the day.

The epoch can be either 1900 or 1904. Excel for Windows uses 1900 and Excel for Macintosh uses 1904. The epochs are:

    1900: 0 January 1900 i.e. 31 December 1899
    1904: 1 January 1904

By default Spreadsheet::WriteExcel uses the Windows/1900 format although it generally isn't an issue since Excel on Windows and the Macintosh will convert automatically between one system and the other. To use the 1904 epoch you must use the C<set_1904()> workbook method.

There are two things to note about the 1900 date format. The first is that the epoch starts on 0 January 1900. The second is that the year 1900 is erroneously but deliberately treated as a leap year. Therefore you must add an extra day to dates after 28 February 1900. The reason for this anomaly is explained at http://support.microsoft.com/support/kb/articles/Q181/3/70.asp 

A date or time in Excel is like any other number. To display the number as a date you must apply a number format to it. Refer to the C<set_num_format()> method above:

    $format->set_num_format('mmm d yyyy hh:mm AM/PM');
    $worksheet->write('A1', 36892.521 , $format); # Jan 1 2001 12:30 AM


The C<Spreadsheet::WriteExcel::Utility> module that is included in the distro contains helper functions for dealing with dates and times in Excel, for example:

    $date = xl_date_list(2002, 1, 1);         # 37257
    $date = xl_parse_date("11 July 1997");    # 35622
    $time = xl_parse_time('3:21:36 PM');      # 0.64
    $date = xl_decode_date_EU("13 May 2002"); # 37389

These functions deal automatically with the s1900 leap year issue described above.

The date and time functions are based on functions provided by the C<Date::Calc> and C<Date::Manip> modules. These modules are very useful if you plan to manipulate dates in different formats. 

There is also the C<excel_date1.pl> program in the C<examples> directory of the WriteExcel distribution which was written by Andrew Benham. It contains a detailed description of the problems involved in calculating dates in Excel. It does not require any external modules.

It is also possible to get Excel to calculate dates for you by defining a function:

    $worksheet->write('A1', '=DATEVALUE("1-Jan-2001")');

However, this carries a performance overhead in Spreadsheet::WriteExcel due to the parsing of the formula and it shouldn't be used for programs that deal with a large number of dates.




=head1 FORMULAS AND FUNCTIONS IN EXCEL

The first thing to note is that there are still some outstanding issues with the implementation of formulas and functions:

    * Writing a formula is much slower than writing the equivalent string.
    * Unary minus isn't supported.
    * You cannot use arrays constants, i.e. {1;2;3}, in functions.
    * You cannot use embedded double quotes in strings.
    * Whitespace is not preserved around operators.

However, these constraints will be removed in future versions. They are here because of a trade-off between features and time.

The following is a brief introduction to formulas and functions in Excel and Spreadsheet::WriteExcel.

A formula is a string that begins with an equals sign:

    '=A1+B1'
    '=AVERAGE(1, 2, 3)'

The formula can contain numbers, strings, boolean values, cell references, cell ranges and functions. Formulas should be written as they appear in Excel, that is cells and functions must be in uppercase.

Cells in Excel are referenced using the A1 notation system where the column is designated by a letter and the row by a number. Columns range from A to IV i.e. 0 to 255, rows range from 1 to 16384. The C<Spreadsheet::WriteExcel::Utility> module that is included in the distro contains helper functions for dealing with A1 notation, for example:

    use Spreadsheet::WriteExcel::Utility;
    
    ($row, $col) = xl_cell_to_rowcol('C2');  # (1, 2)
    $str         = xl_rowcol_to_cell(1, 2);  # C2

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

The sheet reference and the cell reference are separated by  C<!> the exclamation mark symbol. If worksheet names contain spaces then Excel requires that the name is enclosed in single quotes as shown in the last two examples above. In this case you will have to use the quote operator C<q{}> to protect the quotes. See C<perlop> in the main Perl documentation. Only valid sheet names that have been added using the C<addworksheet()> method can be used in formulas. You cannot reference external workbooks.


The following table lists the operators that are available in Excel's formulas. The majority of the operators are the same as Perl's, differences are indicated:

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

The range and comma operators can have different symbols in non-English versions of Excel. These will be supported in a later version of Spreadsheet::WriteExcel. European users of Excel take note:

    $worksheet->write('A1', '=SUM(1; 2; 3)'); # Wrong!!
    $worksheet->write('A1', '=SUM(1, 2, 3)'); # Okay

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
    4. Ensure that you are using ':' as the range operator, A1:A4.
    5. Ensure that you are using ',' as the union operator, SUM(1,2,3).
    6. Ensure the function is in the above table.

If you go through steps 1-6 and you still have a problem, mail me.




=head1 EXAMPLES




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
    foreach my $worksheet ($workbook->sheets()) {
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


=head2 Additional Examples

If you performed a normal installation the following examples files should have been copied to your C<~site/Spreadsheet/WriteExcel/examples> directory:

The following is a description of the example files that are provided
with Spreadsheet::WriteExcel. They are intended to demonstrate the
different features and options of the module.


    Getting started
    ===============
    simple.pl           An example of some of the basic features.
    regions.pl          Demonstrates multiple worksheets.
    stats.pl            Basic formulas and functions.
    formats.pl          Creates a demo of the available formatting.
    demo.pl             Creates a demo of some of the features.


    Advanced
    ========
    sales.pl            An example of a simple sales spreadsheet.
    stocks.pl           Demonstrates conditional formatting.
    write_array.pl      Example of writing 1D or 2D arrays of data.
    chess.pl            An example of formatting using properties.
    images.pl           Adding bitmap images to worksheets.
    stats_ext.pl        Same as stats.pl with external references.
    cgi.pl              A simple CGI program.
    mod_perl.pl         A simple mod_perl program.
    hyperlink1.pl       Shows how to create web hyperlinks.
    hyperlink2.pl       Examples of internal and external hyperlinks.
    merge1.pl           A simple example of cell merging.
    merge2.pl           A more advanced example of cell merging.
    merge3.pl           Merge hyperlinks and merge vertically.
    textwrap.pl         Demonstrates text wrapping options.
    panes.pl            An examples of how to create panes.
    protection.pl       Example of cell locking and formula hiding.
    copyformat.pl       Example of copying a cell format.
    win32ole.pl         A sample Win32::OLE example for comparison.
    easter_egg.pl       Expose the Excel97 flight simulator. A must see.


    Utility
    =======
    convertA1.pl        Helper functions for dealing with A1 notation.
    lecxe.pl            Convert Excel to WriteExcel using Win32::OLE.
    csv2xls.pl          Program to convert a CSV file to an Excel file.
    tab2xls.pl          Program to convert a tab separated file to xls.
    datecalc1.pl        Convert Unix/Perl time to Excel time.
    datecalc2.pl        Calculate an Excel date using Date::Calc.
    writemany.pl        Write an 2d array of values in one go.


    Developer
    =========
    function_locale.pl  Add non-English function names to Formula.pm.
    filehandle.pl       Examples of working with filehandles.
    writeA1.pl          Example of how to extend the module.
    comments.pl         Add cell comments to Excel 5 worksheets.
    bigfile.pl          Write past the 7MB limit with OLE::Storage_Lite.


There are additional examples of a CGI application that uses Spreadsheet::WriteExcel available at the website of the German Unix/web journal iX:
ftp://ftp.heise.de/pub/ix/ix_listings/2001_06/perl.tgz




=head1 LIMITATIONS

The following limits are imposed by Excel or the version of the BIFF file that has been implemented:

    Description                          Limit   Source
    -----------------------------------  ------  -------
    Maximum number of chars in a string  255     Excel 5
    Maximum number of columns            256     Excel All versions
    Maximum number of rows in Excel 5    16384   Excel 5
    Maximum number of rows in Excel 97   65536   Excel 97
    Maximum chars in a sheet name        31      Excel All versions
    Maximum chars in a header/footer     254     Excel All versions


Note: the maximum row reference in a formula is the Excel 5 row limit of 16384.

The minimum file size is 6K due to the OLE overhead. The maximum file size is approximately 7MB (7087104 bytes) of BIFF data. This can be extended by using Takanori Kawai's OLE::Storage_Lite module http://search.cpan.org/search?dist=OLE-Storage_Lite see the C<bigfile.pl> example in the C<examples> directory of the distro.




=head1 REQUIREMENTS

This module requires Perl 5.005 (or later) and Parse::RecDescent: http://search.cpan.org/search?dist=Parse-RecDescent




=head1 PORTABILITY

Spreadsheet::WriteExcel will work on the majority of Windows, UNIX and Macintosh platforms. Specifically, the module will work on any system where perl packs floats in the 64 bit IEEE format. The float must also be in little-endian format but it will be reversed if necessary. Thus:

    print join(" ", map { sprintf "%#02x", $_ } unpack("C*", pack "d", 1.2345)), "\n";

should give (or in reverse order):

    0x8d 0x97 0x6e 0x12 0x83 0xc0 0xf3 0x3f

In general, if you don't know whether your system supports a 64 bit IEEE float or not, it probably does. If your system doesn't, WriteExcel will C<croak()> with the message given in the L<DIAGNOSTICS> section. You can check which platforms the module has been tested on at the CPAN testers site: http://testers.cpan.org/search?request=dist&dist=Spreadsheet-WriteExcel




=head1 DIAGNOSTICS


=over 4

=item Filename required by Spreadsheet::WriteExcel->new()

A filename must be given in the constructor.

=item Can't open filename. It may be in use or protected.

The file cannot be opened for writing. The directory that you are writing to  may be protected or the file may be in use by another program.

=item Unable to create tmp files via IO::File->new_tmpfile().

This is a C<-w> warning. You will see it if you are using Spreadsheet::WriteExcel in an environment where temporary files cannot be created, in which case all data will be stored in memory. The warning is for information only: it does not affect execution but it may affect the speed of execution for large files.

=item Maximum file size, 7087104, exceeded.

The current OLE implementation only supports a maximum BIFF file of this size. This limit can be extended, see the L<LIMITATIONS> section.

=item Can't locate Parse/RecDescent.pm in @INC ...

Spreadsheet::WriteExcel requires the Parse::RecDescent module. Download it from CPAN: http://search.cpan.org/search?dist=Parse-RecDescent

=item Couldn't parse formula ...

There are a large number of warnings which relate to badly formed formulas and functions. See the L<FORMULAS AND FUNCTIONS IN EXCEL> section for suggestions on how to avoid these errors.

=item Required floating point format not supported on this platform.

Operating system doesn't support 64 bit IEEE float or it is byte-ordered in a way unknown to WriteExcel.

=back




=head1 THE EXCEL BINARY FORMAT

The following is some general information about the Excel binary format for anyone who may be interested.

Excel data is stored in the "Binary Interchange File Format" (BIFF) file format. Details of this format are given in the Excel SDK, the "Excel Developer's Kit" from Microsoft Press. It is also included in the MSDN CD library but is no longer available on the MSDN website. An older version of the BIFF documentation is available at http://www.cubic.org/source/archive/fileform/misc/excel.txt

Charles Wybble has collected together almost all of the available information about the Excel file format. See "The Chicago Project" at http://chicago.sourceforge.net/devel/

Daniel Rentz of OpenOffice has also written a detailed description of the Excel workbook records, see http://sc.openoffice.org/excelfileformat.pdf

The BIFF portion of the Excel file is comprised of contiguous binary records that have different functions and that hold different types of data. Each BIFF record is comprised of the following three parts:

        Record name;   Hex identifier, length = 2 bytes
        Record length; Length of following data, length = 2 bytes
        Record data;   Data, length = variable

The BIFF data is stored along with other data in an OLE Compound File. This is a structured storage which acts like a file system within a file. A Compound File is comprised of storages and streams which, to follow the file system analogy, are like directories and files.

The documentation for the OLE::Storage module, http://user.cs.tu-berlin.de/~schwartz/pmh/guide.html , contains one of the few descriptions of the OLE Compound File in the public domain. The Digital Imaging Group have also detailed the OLE format in the JPEG2000 specification: see Appendix A of http://www.i3a.org/pdf/wg1n1017.pdf

For a open source implementation of the OLE library see the 'cole' library at http://atena.com/libole2.php

The source code for the Excel plugin of the Gnumeric spreadsheet also contains information relevant to the Excel BIFF format and the OLE container, http://www.ximian.com/apps/gnumeric.php3 and ftp://ftp.ximian.com/pub/ximian-source/

In addition the source code for OpenOffice is available at http://www.openoffice.org/

An article describing Spreadsheet::WriteExcel and how it works appears in Issue #19 of The Perl Journal, http://www.samag.com/documents/s=1272/sam05030004/ It is reproduced, by kind permission, in the C<doc> directory of the distro.


Please note that the provision of this information does not constitute an invitation to start hacking at the BIFF or OLE file formats. There are more interesting ways to waste your time. ;-)




=head1 WRITING EXCEL FILES

Depending on your requirements, background and general sensibilities you may prefer one of the following methods of getting data into Excel:

* Win32::OLE module and office automation. This requires a Windows platform and an installed copy of Excel. This is the most powerful and complete method for interfacing with Excel. See http://www.activestate.com/ASPN/Reference/Products/ActivePerl-5.6/faq/Windows/ActivePerl-Winfaq12.html and http://www.activestate.com/ASPN/Reference/Products/ActivePerl-5.6/site/lib/Win32/OLE.html If your main platform is UNIX but you have the resources to set up a separate Win32/MSOffice server, you can convert office documents to text, postscript or PDF using Win32::OLE. For a demonstration of how to do this using Perl see Docserver: http://search.cpan.org/search?mode=module&query=docserver

* CSV, comma separated variables or text. If the file extension is C<csv>, Excel will open and convert this format automatically. Generating a valid CSV file isn't as easy as it seems. Have a look at the DBD::RAM, DBD::CSV, Text::xSV and Text::CSV_XS modules.

* DBI with DBD::ADO or DBD::ODBC. Excel files contain an internal index table that allows them to act like a database file. Using one of the standard Perl database modules you can connect to an Excel file as a database.

* DBD::Excel, you can also access Spreadsheet::WriteExcel using the standard DBI interface via Takanori Kawai's DBD::Excel module http://search.cpan.org/search?dist=DBD-Excel.

* Spreadsheet::WriteExcel::Simple for an easier interface to a new Excel file: http://search.cpan.org/search?dist=Spreadsheet-WriteExcel-Simple

* Spreadsheet::WriteExcel::FromDB to create an Excel file directly from a DB table: http://search.cpan.org/search?dist=Spreadsheet-WriteExcel-FromDB

* HTML tables. This is an easy way of adding formatting via a text based format.

* XML, the Excel XML and HTML file specification are available from http://msdn.microsoft.com/library/officedev/ofxml2k/ofxml2k.htm

For other Perl-Excel modules try the following search: http://search.cpan.org/search?mode=module&query=excel




=head1 READING EXCEL FILES

To read data from Excel files try:

* Spreadsheet::ParseExcel. This uses the OLE::Storage-Lite module to extract data from an Excel file. http://search.cpan.org/search?dist=Spreadsheet-ParseExcel

* Spreadsheet::ParseExcel_XLHTML. This module uses Spreadsheet::ParseExcel's interface but uses xlHtml (see below) to do the conversion: http://search.cpan.org/search?dist=Spreadsheet-ParseExcel_XLHTML
Spreadsheet::ParseExcel_XLHTML 

* There are also open source C/C++ projects. Try the xlHtml "Excel to HTML Converter" project at http://www.xlhtml.org/ and the OLE Filters Project at http://atena.com/libole2.php. 

* DBD::Excel, you can also access Spreadsheet::ParseExcel using the standard DBI interface via  Takanori Kawai's DBD::Excel module http://search.cpan.org/search?dist=DBD-Excel.

* Win32::OLE module and office automation. See, the section L<WRITING EXCEL FILES>.

* HTML tables. If the files are saved from Excel in a HTML format the data can be accessed using HTML::TableExtract http://search.cpan.org/search?dist=HTML-TableExtract

* DBI with DBD::ADO or DBD::ODBC. See, the section L<WRITING EXCEL FILES>.

* XML::Excel converts Excel files to XML using Spreadsheet::ParseExcel http://search.cpan.org/search?dist=XML-Excel. 

* OLE::Storage, aka LAOLA. This is a Perl interface to OLE file formats. In particular, the distro contains an Excel to HTML converter called Herbert, http://user.cs.tu-berlin.de/~schwartz/pmh/ This has been superseded by the Spreadsheet::ParseExcel module.

For other Perl-Excel modules try the following search: http://search.cpan.org/search?mode=module&query=excel

If you wish to view Excel files on a UNIX/Linux platform check out the excellent Gnumeric spreadsheet application at http://www.gnome.org/projects/gnumeric/ or OpenOffice at http://www.openoffice.org/

If you wish to view Excel files on a Windows platform which doesn't have Excel installed you can use the free Microsoft Excel Viewer http://officeupdate.microsoft.com/2000/downloaddetails/xlviewer.htm





=head1 BUGS

Orange isn't.

Formulas are formulae.

Nested formulas sometimes aren't parsed correctly and give a result of "#VALUE". This will be fixed in a later release.

Spreadsheet::ParseExcel: All formulas created by Spreadsheet::WriteExcel are read as having a value of zero. This is because Spreadsheet::WriteExcel only stores the formula and not the calculated result.

OpenOffice: Numerical formats are not displayed due to some missing records in Spreadsheet::WriteExcel. URLs are not displayed as links.

Gnumeric: Some formatting is not displayed correctly. URLs are not displayed as links.

MS Access: The Excel files that are produced by this module are not compatible with MS Access. Use DBI or ODBC instead.

The lack of a portable way of writing a little-endian 64 bit IEEE float.



=head1 TO DO

The roadmap is as follows:

=over

=item * Move to Excel97/2000 format as standard. This will allow strings greater than 255 characters and hopefully Unicode. The Excel 5 format will be optional. This will be in the next major release of the module. All other features are on hold.

=back

You can keep up to date with future release by registering as a user with Freshmeat http://freshmeat.net/ and subscribing to Spreadsheet::WriteExcel at the project page http://freshmeat.net/projects/writeexcel/ You will then receive mailed updates when a new version is released. Alternatively you can keep an eye on news://comp.lang.perl.announce

Also, here are some of the most requested features that probably won't get added:

=over

=item * Graphs. The format is documented but it would require too much work to implement. It would also require too much work to design a useable interface to the hundreds of features in an Excel graph. So that's two too much works.

=item * Macros. This would solve the previous problem neatly. However, the format of Excel macros isn't documented.

=item * Some feature that you really need. ;-)

=back

If there is some feature of an Excel file that you really, really need then you should use Win32::OLE with Excel on Windows.





=head1 SEE ALSO

Spreadsheet::ParseExcel. http://search.cpan.org/search?dist=Spreadsheet-ParseExcel

Spreadsheet::WriteExcel::Simple. http://search.cpan.org/search?dist=Spreadsheet-WriteExcel-Simple

Spreadsheet::WriteExcel::FromDB. http://search.cpan.org/search?dist=Spreadsheet-WriteExcel-FromDB

"Reading and writing Excel files with Perl" by Teodor Zlatanov, atIBM developerWorks: http://www-106.ibm.com/developerworks/library/l-pexcel/

"Excel-Dateien mit Perl erstellen - Controller im Glück" by Peter Dintelmann and Christian Kirsch in the German Unix/web journal iX: http://www.heise.de/ix/artikel/2001/06/175/

"Spreadsheet::WriteExcel" in The Perl Journal: http://www.samag.com/documents/s=1272/sam05030004/

Spreadsheet::WriteExcel documentation in Japanese by Takanori Kawai. http://member.nifty.ne.jp/hippo2000/perltips/Spreadsheet/WriteExcel.htm

Oesterly user brushes with fame:
http://oesterly.com/releases/12102000.html


=head1 ACKNOWLEDGEMENTS


The following people contributed to the debugging and testing of Spreadsheet::WriteExcel:

Alexander Farber, Arthur@ais, Artur Silveira da Cunha, Borgar Olsen, Brian White, Cedric Bouvier, CPAN testers, Daniel Berger, Daniel Gardner, Ernesto Baschny, Felipe Pérez Galiana, Hanc Pavel, Harold Bamford, James Holmes, Johan Ekenberg, J.C. Wren, Kenneth Stacey, Keith Miller, Kyle Krom, Markus Schmitz, Michael Buschauer, Mike Blazer, Michael Erickson, Paul J. Falbe, Paul Medynski, Peter Dintelmann, Reto Badertscher, Rich Sorden, Shane Ashby, Shenyu Zheng, Steve Sapovits, Sven Passig, Vahe Sarkissian.

The following people contributed code, examples or Excel information:

Andrew Benham, Bill Young, Cedric Bouvier, Daniel Rentz, Ian Penman, Pierre-Jean Vouette, Marco Geri, Sam Kington, Takanori Kawai, Tom O'Sullivan.

Additional thanks to Takanori Kawai for translating the documentation into Japanese.

Dirk Eddelbuettel maintains the Debian distro.

Thanks to Damian Conway for the excellent Parse::RecDescent.

Thanks to Michael Meeks and Jody Goldberg for their work on Gnumeric.




=head1 AUTHOR

John McNamara jmcnamara@cpan.org

    Pointy birds, O pointy pointy,
    Anoint my head, anointy nointy.
    
        -- Steve Martin.


=head1 COPYRIGHT

© MM-MMII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

