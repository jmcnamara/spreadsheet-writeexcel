#!/usr/bin/perl -w

use ExtUtils::MakeMaker;
use strict;

# perl >= 5.005 is required for qr() support in Parse::RecDescent.
# There are also some 4 arg substr()'s in WriteExcel.
#
require 5.005;

my %deps = (
    'Parse::RecDescent' => 0,
    'File::Temp'        => 0,
    'OLE::Storage_Lite' => 0.19,
);

my %resources = (
    repository  => 'http://github.com/jmcnamara/spreadsheet-writeexcel',
    MailingList => 'http://groups.google.com/group/spreadsheet-writeexcel',
);

my %params = (
    'AUTHOR'        => 'John McNamara (jmcnamara@cpan.org)',
    'ABSTRACT'      => 'Write to a cross platform Excel binary file',
    'NAME'          => 'Spreadsheet::WriteExcel',
    'VERSION_FROM'  => 'lib/Spreadsheet/WriteExcel.pm',
    'NEEDS_LINKING' => 0,
    'dist'          => { COMPRESS => 'gzip --best', SUFFIX => 'gz' },
    'EXE_FILES'     => ['bin/chartex'],
    'LICENSE'       => 'perl',
    'PREREQ_PM'     => \%deps,
);

if ( $ExtUtils::MakeMaker::VERSION ge '6.46' ) {
    $params{META_MERGE} = { resources => \%resources };
}

WriteMakefile(%params);

sub MY::postamble {
    "mydocs:\n"
      . "\tperl examples/gen_examples_pod.pl examples "
      . "> lib/Spreadsheet/WriteExcel/Examples.pm\n"
      . "\tpod2cpanhtml lib/Spreadsheet/WriteExcel.pm "
      . qq{ | perl -ne 'print unless /\\s+under Perl v[\\d.]+\\s/' }
      . "> docs/WriteExcel.html";
}
