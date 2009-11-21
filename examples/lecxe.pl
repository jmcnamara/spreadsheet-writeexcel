#!/usr/bin/perl -w

#
# Utility program to convert an Excel file into a Spreadsheet::WriteExcel
# program using Win32::OLE
#

#
# lecxe program
# by t0mas@netlords.net
#
# Version  0.01a    Initial release (alpha)


# Modules
use strict;
use Win32::OLE;
use Win32::OLE::Const;
use Getopt::Std;


# Vars
use vars qw(%opts);


# Get options
getopts('i:o:v',\%opts);


# Not enough options
exit &usage unless ($opts{i} && $opts{o});


# Create Excel object
my $Excel = new Win32::OLE("Excel.Application","Quit") or
        die "Can't start excel: $!";


# Get constants
my $ExcelConst=Win32::OLE::Const->Load("Microsoft Excel");


# Show Excel
$Excel->{Visible} = 1 if ($opts{v});


# Open infile
my $Workbook = $Excel->Workbooks->Open({Filename=>$opts{i}});


# Open outfile
open (OUTFILE,">$opts{o}") or die "Can't open outfile $opts{o}: $!";


# Print header for outfile
print OUTFILE <<'EOH';
#!/usr/bin/perl -w


use strict;
use Spreadsheet::WriteExcel;


use vars qw($workbook %worksheets %formats);


$workbook = Spreadsheet::WriteExcel->new("_change_me_.xls");


EOH


# Loop all sheets
foreach my $sheetnum (1..$Excel->Workbooks(1)->Worksheets->Count) {


        # Format sheet
        my $name=$Excel->Workbooks(1)->Worksheets($sheetnum)->Name;
        print "Sheet $name\n" if ($opts{v});
        print OUTFILE "# Sheet $name\n";
        print OUTFILE "\$worksheets{'$name'} = \$workbook->add_worksheet('$name');\n";


        # Get usedrange of cells in worksheet
        my $usedrange=$Excel->Workbooks(1)->Worksheets($sheetnum)->UsedRange;


        # Loop all columns in used range
        foreach my $j (1..$usedrange->Columns->Count){


                # Format column
                print "Col $j\n" if ($opts{v});
                my ($colwidth);
                $colwidth=$usedrange->Columns($j)->ColumnWidth;
                print OUTFILE "# Column $j\n";
                print OUTFILE "\$worksheets{'$name'}->set_column(".($j-1).",".($j-1).
                        ", $colwidth);\n";


                # Loop all rows in used range
                foreach my $i (1..$usedrange->Rows->Count){


                        # Format row
                        print "Row $i\n" if ($opts{v});
                        print OUTFILE "# Row $i\n";
                        do {
                                my ($rowheight);
                                $rowheight=$usedrange->Rows($i)->RowHeight;
                                print OUTFILE "\$worksheets{'$name'}->set_row(".($i-1).
                                        ", $rowheight);\n";
                        } if ($j==1);


                        # Start creating cell format
                        my $fname="\$formats{'".$name.'R'.$i.'C'.$j."'}";
                        my $format="$fname=\$workbook->add_format();\n";
                        my $print_format=0;

                        # Check for borders
                        my @bfnames=qw(left right top bottom);
                        foreach my $k (1..$usedrange->Cells($i,$j)->Borders->Count) {
                                my $lstyle=$usedrange->Cells($i,$j)->Borders($k)->LineStyle;
                                if ($lstyle > 0) {
                                        $format.=$fname."->set_".$bfnames[$k-1]."($lstyle);\n";
                                        $print_format=1;
                                }
                        }


                        # Check for font
                        my ($fontattr,$prop,$func,%fontsets,$fontColor);
                        %fontsets=(Name=>'set_font',
                                                Size=>'set_size');
                        while (($prop,$func) = each %fontsets) {
                                $fontattr=$usedrange->Cells($i,$j)->Font->$prop;
                                if ($fontattr ne "") {
                                        $format.=$fname."->$func('$fontattr');\n";
                                        $print_format=1;
                                }


                        }
                        %fontsets=(Bold=>'set_bold(1)',
                                                Italic=>'set_italic(1)',
                                                Underline=>'set_underline(1)',
                                                Strikethrough=>'set_strikeout(1)',
                                                Superscript=>'set_script(1)',
                                                Subscript=>'set_script(2)',
                                                OutlineFont=>'set_outline(1)',
                                                Shadow=>'set_shadow(1)');
                        while (($prop,$func) = each %fontsets) {
                                $fontattr=$usedrange->Cells($i,$j)->Font->$prop;
                                if ($fontattr==1) {
                                        $format.=$fname."->$func;\n" ;

                                        $print_format=1;
                                }
                        }
                        $fontColor=$usedrange->Cells($i,$j)->Font->ColorIndex();
                        if ($fontColor>0&&$fontColor!=$ExcelConst->{xlColorIndexAutomatic}) {
                                $format.=$fname."->set_color(".($fontColor+7).");\n" ;
                                $print_format=1;
                        }



                        # Check text alignment, merging and wrapping
                        my ($halign,$valign,$merge,$wrap);
                        $halign=$usedrange->Cells($i,$j)->HorizontalAlignment;
                        my %hAligns=($ExcelConst->{xlHAlignCenter}=>"'center'",
                                $ExcelConst->{xlHAlignJustify}=>"'justify'",
                                $ExcelConst->{xlHAlignLeft}=>"'left'",
                                $ExcelConst->{xlHAlignRight}=>"'right'",
                                $ExcelConst->{xlHAlignFill}=>"'fill'",
                                $ExcelConst->{xlHAlignCenterAcrossSelection}=>"'merge'");
                        if ($halign!=$ExcelConst->{xlHAlignGeneral}) {
                                $format.=$fname."->set_align($hAligns{$halign});\n";
                                $print_format=1;
                        }
                        $valign=$usedrange->Cells($i,$j)->VerticalAlignment;
                        my %vAligns=($ExcelConst->{xlVAlignBottom}=>"'bottom'",
                                $ExcelConst->{xlVAlignCenter}=>"'vcenter'",
                                $ExcelConst->{xlVAlignJustify}=>"'vjustify'",
                                $ExcelConst->{xlVAlignTop}=>"'top'");
                        if ($valign) {
                                $format.=$fname."->set_align($vAligns{$valign});\n";
                                $print_format=1;
                        }
                        $merge=$usedrange->Cells($i,$j)->MergeCells;
                        if ($merge==1) {
                                $format.=$fname."->set_merge();\n";

                                $print_format=1;
                        }
                        $wrap=$usedrange->Cells($i,$j)->WrapText;
                        if ($wrap==1) {
                                $format.=$fname."->set_text_wrap(1);\n";

                                $print_format=1;
                        }


                        # Check patterns
                        my ($pattern,%pats);
                        %pats=(-4142=>0,-4125=>2,-4126=>3,-4124=>4,-4128=>5,-4166=>6,
                                        -4121=>7,-4162=>8);
                        $pattern=$usedrange->Cells($i,$j)->Interior->Pattern;
                        if ($pattern&&$pattern!=$ExcelConst->{xlPatternAutomatic}) {
                                $pattern=$pats{$pattern} if ($pattern<0 && defined $pats{$pattern});
                                $format.=$fname."->set_pattern($pattern);\n";

                                # Colors fg/bg
                                my ($cIndex);
                                $cIndex=$usedrange->Cells($i,$j)->Interior->PatternColorIndex;
                                if ($cIndex>0&&$cIndex!=$ExcelConst->{xlColorIndexAutomatic}) {
                                        $format.=$fname."->set_bg_color(".($cIndex+7).");\n";
                                }
                                $cIndex=$usedrange->Cells($i,$j)->Interior->ColorIndex;
                                if ($cIndex>0&&$cIndex!=$ExcelConst->{xlColorIndexAutomatic}) {
                                        $format.=$fname."->set_fg_color(".($cIndex+7).");\n";
                                }
                                $print_format=1;
                        }


                        # Check for number format
                        my ($num_format);
                        $num_format=$usedrange->Cells($i,$j)->NumberFormat;
                        if ($num_format ne "") {
                                $format.=$fname."->set_num_format('$num_format');\n";
                                $print_format=1;
                        }


                        # Check for contents (text or formula)
                        my ($contents);
                        $contents=$usedrange->Cells($i,$j)->Formula;
                        $contents=$usedrange->Cells($i,$j)->Text if ($contents eq "");


                        # Print cell
                        if ($contents ne "" or $print_format) {
                                print OUTFILE "# Cell($i,$j)\n";
                                print OUTFILE $format if ($print_format);
                                print OUTFILE "\$worksheets{'$name'}->write(".($i-1).",".($j-1).
                                        ",'$contents'";
                                print OUTFILE ",$fname" if ($print_format);
                                print OUTFILE ");\n";
                        }
                }
        }
}


# Famous last words...
print OUTFILE "\$workbook->close();\n";


# Close outfile
close (OUTFILE) or die "Can't close outfile $opts{o}: $!";


####################################################################
sub usage {
        printf STDERR "usage: $0 [options]\n".
                "\tOptions:\n".
                "\t\t-v       \tverbose mode\n" .
                "\t\t-i <name>\tname of input file\n" .
                "\t\t-o <name>\tname of output file\n";
}


####################################################################
sub END {
        # Quit excel
        do {
                $Excel->{DisplayAlerts} = 0;
                $Excel->Quit;
        } if (defined $Excel);
}


__END__


=head1 NAME


lecxe - A Excel file to Spreadsheet::WriteExcel code converter


=head1 DESCRIPTION


This program takes an MS Excel workbook file as input and from
that file, produces an output file with Perl code that uses the
Spreadsheet::WriteExcel module to reproduce the original
file.


=head1 STUFF


Additional hands-on editing of the output file might be neccecary
as:


* This program always names the file produced by output script
  _change_me_.xls


* Users of international Excel versions will have som work to do
  on list separators and numeric punctation characters.


=head1 SEE ALSO


L<Win32::OLE>, L<Win32::OLE::Variant>, L<Spreadsheet::WriteExcel>


=head1 BUGS


* Picks wrong color on cells sometimes.


* Probably a few other...


=head1 DISCLAIMER


I do not guarantee B<ANYTHING> with this program. If you use it you
are doing so B<AT YOUR OWN RISK>! I may or may not support this
depending on my time schedule...


=head1 AUTHOR


t0mas@netlords.net


=head1 COPYRIGHT


Copyright 2001, t0mas@netlords.net


This package is free software; you can redistribute it and/or
modify it under the same terms as Perl itself.
