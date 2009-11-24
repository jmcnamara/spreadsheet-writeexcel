#!/usr/bin/perl

###############################################################################
#
# Simple utility to convert the example programs listed in the README file into
# a Pod doc for easier access via CPAN.
#
# reverse('©'), November 2009, John McNamara, jmcnamara@cpan.org
#

use 5.008;
use strict;
use warnings;

main();

###############################################################################
#
# main()
#
# Convert the example programs listed in the README file into a Pod doc for
# easier access via CPAN.
#
sub main {

    my @examples;
    my $examples_dir = $ARGV[0] || '.';

    # Get the version from the local WriteExcel.pm.
    require "$examples_dir/../lib/Spreadsheet/WriteExcel.pm";
    my $version = Spreadsheet::WriteExcel->VERSION();

    # Read the filenames and descriptions from the examples README file.
    open my $readme, '<', $examples_dir . '/README'
      or die "Couldn't open $examples_dir/README file: $!\n";

    while ( my $line = <$readme> ) {
        if ( $line =~ /^\w+.pl\s/ ) {
            chomp $line;
            my ( $filename, $description ) = split " ", $line, 2;
            push @examples, [ $filename, $description ];
        }
    }

    die "Didn't find example programs in $examples_dir/README\n"
      unless @examples;

    print_header($version);
    print_index(@examples);

    for my $example (@examples) {
        my $filename = $example->[0];
        print_example( $examples_dir, $filename, $version );
    }

    print_footer();
}

###############################################################################
#
# print_header()
#
# Print the header section of the Pod documentation.
#
sub print_header {

    my $version = shift;

    # I just don't like here docs.
    print "package Spreadsheet::WriteExcel::Examples;\n\n";

    print '#' x 79, "\n";
    print "#\n";
    print "# Examples - Spreadsheet::WriteExcel examples.\n";
    print "#\n";

    print "# A documentation only module showing the examples that are\n";
    print "# included in the Spreadsheet::WriteExcel distribution. This\n";
    print "# file was generated automatically via the gen_examples_pod.pl\n";
    print "# program that is also included in the examples directory.\n";
    print "#\n";

    print "# Copyright 2000-2009, John McNamara, jmcnamara\@cpan.org\n";
    print "#\n";
    print "# Documentation after __END__\n";
    print "#\n\n";

    print "use strict;\n";
    print "use vars qw(\$VERSION);\n";
    print "\$VERSION = '$version';\n\n";

    print "1;\n";
    print "\n";
    print "__END__\n\n";

    print "=pod\n\n";

    print "=head1 NAME\n\n";

    print "Examples - Spreadsheet::WriteExcel example programs.\n\n";

    print "=head1 DESCRIPTION\n\n";

    print "This is a documentation only module showing the examples that are\n";
    print "included in the L<Spreadsheet::WriteExcel> distribution.\n\n";
    print "This file was auto-generated via the gen_examples_pod.pl\n";
    print "program which is shown below.\n";
    print "\n";

}

###############################################################################
#
# print_index()
#
# Print an index to the example programs with the short descriptions from the
# README file and a link to the appropriate section.
#
sub print_index {

    my @examples = @_;
    my $count    = scalar @examples;

    print "=head1 Example programs\n\n";

    print "The following is a list of the $count example programs that are ";
    print "included in the Spreadsheet::WriteExcel distribution.\n\n";

    print "=over\n\n";

    for my $example (@examples) {
        print "=item * L<Example: ";
        print $example->[0];
        print "> ";
        print $example->[1];
        print "\n\n";
    }

    print "=back\n\n";
}

###############################################################################
#
# print_example()
#
# Print each example program in its own =head1 section with a short description
# extracted from the first comment section of at the start and the code
# in a Pod verbatim section.
#
sub print_example {

    my $examples_dir = shift;
    my $example      = shift;
    my $version      = shift;
    my $verbatim     = '';
    my $in_header    = 0;

    open my $example_fh, '<', $examples_dir . '/' . $example;

    if ( !defined $example_fh ) {
        warn "Couldn't open $examples_dir/$example: $!\n";
        return undef;
    }

    print "=head2 Example: $example\n\n";

    while ( my $line = <$example_fh> ) {
        $line =~ s/\r//;
        $verbatim .= '    ' . $line;

        # Ignore the most common copyright line.
        next if $line =~ m/jmcnamara/;

        # Look for the first comment section but ignore the #!perl shebang line.
        if ( $in_header == 0 && $line !~ m/perl/ && $line =~ m/^#/ ) {
            $in_header = 1;
        }

        # In the first comment section.
        if ( $in_header == 1 ) {

            # Unset flag when leaving the first comment section.
            $in_header++ if $line !~ m/^#/;

            # Remove the comment char and the first leading space. This maintain
            # any embedded verbatim like sections.
            $line =~ s/^#+[ ]{0,1}//;

            print $line;
        }
    }

    print $verbatim, "\n\n";

    print 'Download this example: L<http://cpansearch.perl.org/src/JMCNAMARA/';
    print "Spreadsheet-WriteExcel-$version/examples/$example>\n\n";
}

###############################################################################
#
# print_footer()
#
# Print the footer section of the Pod documentation
#
sub print_footer {

    print "=head1 AUTHOR\n\n";

    print "John McNamara jmcnamara\@cpan.org\n\n";

    print "Contributed examples contain the original author's name.\n\n";

    print "=head1 COPYRIGHT\n\n";

    print "Copyright MM-MMIX, John McNamara.\n\n";

    print "All Rights Reserved. This module is free software. It may be used, ";
    print "redistributed and/or modified under the same terms as Perl itself.";
    print "\n\n";

    print "=cut\n";

}

__END__
