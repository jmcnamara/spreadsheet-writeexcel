package Spreadsheet::WriteExcel;

######################################################################
#
# WriteExcel.
#
# Spreadsheet::WriteExcel - Write text and numbers to a cross-platform
# Excel binary file.
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

$VERSION = '0.21'; # 1 October 2000, Carlos Williams correct

######################################################################
#
# new()
#
# Constructor. Wrapper for a Workbook object.
# uses: Spreadsheet::BIFFwriter
#       Spreadsheet::OLEwriter
#       Spreadsheet::Workbook
#       Spreadsheet::Worksheet
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

Spreadsheet::WriteExcel - Write text and numbers to a cross-platform Excel binary file.




=head1 VERSION

This document refers to version 0.21 of Spreadsheet::WriteExcel, released October 1, 2000.




=head1 SYNOPSIS

To write a string and a number to the first worksheet in an Excel workbook called perl.xls:

    use Spreadsheet::WriteExcel;

    $row1 = $col1 = 0;
    $row2 = 1;

    my $workbook = Spreadsheet::WriteExcel->new("perl.xls");
    $worksheet  = $workbook->addworksheet();

    $worksheet->write($row1, $col1, "Hi Excel!");
    $worksheet->write($row2, $col1, 1.2345);




=head1 DESCRIPTION

The Spreadsheet::WriteExcel module can be used to write numbers and text in the native Excel binary file format. Multiple worksheets can be added to a workbook. Formatting of cells is not yet supported.

The Excel file produced by this module is compatible with Excel 5, 95, 97 and 2000.

The module will work on the majority of Windows, UNIX and Macintosh platforms. Generated files are also compatible with the Linux/UNIX spreadsheet applications Star Office, Gnumeric and XESS. The generated files are not compatible with MS Access. 




=head1 WORKBOOK METHODS

The Spreadsheet::WriteExcel module provides an object oriented interface to a new Excel workbook.The following methods are available through a new workbook.




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




=head2 close()

The C<close()> method can be called to explicitly close an Excel file. Otherwise the file will be closed automatically when the object reference goes out of scope or when the program ends.

    $workbook->close();

In general it is only necessary to explicitly close a file if you want to perform some other operation on it such as copying or checking the size.




=head2 worksheets()

The C<worksheets()> method returns a reference to the array of worksheets in a workbook. This can be useful if you want to repeat an operation on each worksheet in a workbook or where you wish to refer to a worksheet by its index:

    foreach $worksheet (@{$workbook->worksheets()}) {
       $worksheet->write(0, 0, "Hello");
    }
    
    # or:
    
    $worksheets = $workbook->worksheets();
    @$worksheets[0]->write(0, 0, "Hello");


References are explained in detail in C<perlref> and C<perlreftut> in the main Perl documentation.




=head1 WORKSHEET METHODS

The following methods are available through to a new worksheet.




=head2 write($row, $column, $token)

The  C<write()> method calls C<write_number()> if C<$token> matches the following regex:

    $token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/

Otherwise it calls C<write_string()>:

    $worksheet->write(0, 0, "Hello" );  # write_string()
    $worksheet->write(1, 0, "One"   );  # write_string()
    $worksheet->write(2, 0,  2      );  # write_number()
    $worksheet->write(3, 0,  3.00001);  # write_number()

It should be noted that C<$row> and C<$column> are zero indexed cell locations for the C<write> methods. Thus, Cell A1 is (0, 0) and Cell AD2000 is (1999, 29). Cells can be written to in any order. They can also be overwritten.

The C<write> methods return:

    0 for success
   -1 for insufficient number of arguments
   -2 for row or column out of bounds
   -3 for string too long.



=head2 write_number($row, $column, $number)

Write an integer or a float to the cell specified by C<$row> and C<$column>:

    $worksheet->write_number(0, 0,  1     );
    $worksheet->write_number(1, 0,  2.3451);




=head2 write_string($row, $column, $string)

Write a string to the cell specified by C<$row> and C<$column>:

    $worksheet->write_string(0, 0, "Your text here" );

The maximum string size is 255 characters.




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




=head2 set_col_width($first_col, $last_col, $width);

This method can be used to specify the width of a single column or a range of columns. If the method is applied to a single column the value of C<$first_col> and C<$last_col> should be the same:

    $worksheet->set_col_width(0, 0, 20);
    $worksheet->set_col_width(1, 3, 30);

The width corresponds to the column width value that is specified in Excel. It is approximately equal to the length of a string in the default font of Arial 10.




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

    # Add a caption to each worksheet
    foreach my $worksheet (@{$workbook->worksheets()}) {
       $worksheet->write(0, 0, "Sales");
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


In general, if you don't know whether your system supports a 64 bit IEEE float or not, it probably does. If your system doesn't WriteExcel will C<croak()> with the message given in the Diagnostics section.


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

Another useful source is the filters project http://arturo.directmail.org/filtersweb/ 

The source code for the Excel plugin of the Gnumeric spreadsheet also contains information relevant to the Excel BIFF format and the OLE container, http://www.gnumeric.org/

The soon to be GPLed source code for Star Office should also be of interest, http://www.openoffice.org/

Please note that the provision of this information does not constitute an invitation to start hacking at the BIFF or OLE file formats. There are more interesting ways to waste your time. ;)




=head1 WRITING EXCEL FILES

Depending on your requirements, background and general sensibilities you may prefer one of the following methods of getting data into Excel:

* CSV, comma separated variables or text. If the file extension is C<csv>, Excel will open and convert this format automatically.

* DBI, ADO or ODBC. Connect to an Excel file as a database. Using the appropriate driver Excel will behave like a database.

* HTML tables. This is an easy way of adding formatting.

* Win32::OLE module and office automation. See, the section "Reading Excel Files".



=head1 READING EXCEL FILES

Despite the title of this module the most commonly asked questions are in relation to reading Excel files. To read data from Excel files try:

* DBI, ADO or ODBC. Connect to an Excel file as a database. Using the appropriate driver Excel will behave like a database.

* HTML tables. If the files are saved from Excel in a HTML format the data can be accessed using HTML::TableExtract http://search.cpan.org/search?dist=HTML-TableExtract

* OLE::Storage, aka LAOLA. This is a Perl interface to OLE file formats. In particular, the distro contains an Excel to HTML converter called Herbert, http://user.cs.tu-berlin.de/~schwartz/pmh/ There is also an open source C/C++ project based on the LAOLA work. Try the Filters Project at http://arturo.directmail.org/filtersweb/ and the xlHtml Project at http://www.xlhtml.org/ The xlHtml filter is more complete than Herbert.

* Win32::OLE module and office automation. This requires a Windows platform and an installed copy of Excel. This is the most powerful and complete method for interfacing with Excel. See http://www.activestate.com/Products/ActivePerl/docs/faq/Windows/ActivePerl-Winfaq12.html and http://www.activestate.com/Products/ActivePerl/docs/site/lib/Win32/OLE.html

If your main platform is UNIX but you have the resources to set up a separate Win32/MSOffice server, you can convert office documents to text, postscript or PDF using Win32:OLE. For a demonstration of how to do this using Perl see Docserver: http://search.cpan.org/search?mode=module&query=docserver

If you wish to view Excel files on a UNIX/Linux platform check out the excellent Gnumeric spreadsheet application at http://www.gnumeric.org/gnumeric

If you wish to view Excel files on Windows platforms which don't have Excel installed you can use the free Microsoft Excel Viewer http://officeupdate.microsoft.com/downloadDetails/xlviewer.htm




=head1 BUGS

The Excel files that are produced by this module are not compatible with MS Access. Use DBI or ODBC instead.


The lack of a portable way of writing a little-endian 64 bit IEEE float.

QuickView: If you wish to write files are fully compatible with QuickView it is necessary to write the cells in a sequential row by row order.




=head1 TO DO

This module will be extended to include the following, probably in this order:

    1. Cell and font formatting
    2. Row and column formatting
    3. Unlimited file size
    4. Document summary information
    5. Formulas (hopefully)


Items 1. and 2. should be ready be Nov 1 2000 unless you are reading this in December 2000.




=head1 ACKNOWLEDGEMENTS

The following people contributed to the debugging and testing of WriteExcel.pm:

Arthur@ais, Artur Silveira da Cunha, CPAN testers, Daniel Gardner, Harold Bamford, Johan Ekenberg, John Wren, Michael Buschauer, Mike Blazer, Paul J. Falbe.




=head1 AUTHOR

John McNamara jmcnamara@cpan.org

        I have eaten
        the plums
        that were in
        the icebox

        and which
        you were probably
        saving
        for breakfast

        Forgive me
        they were delicious
        so sweet
        and so cold

        - William Carlos Williams




=head1 COPYRIGHT

Copyright (c) 2000, John McNamara. All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
