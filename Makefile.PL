use ExtUtils::MakeMaker;

# perl 5.005 is required for qr() support in Parse::RecDescent.
# There are also some 4 arg substr()'s in WriteExcel.
#
require 5.005;

WriteMakefile(
    'AUTHOR'        => 'John McNamara (jmcnamara@cpan.org)',
    'ABSTRACT'      => 'Write to a cross platform Excel binary file',
    'NAME'          => 'Spreadsheet::WriteExcel',
    'VERSION_FROM'  => 'lib/Spreadsheet/WriteExcel.pm',
    'NEEDS_LINKING' => 0,
    'PREREQ_PM'     => {Parse::RecDescent => 0, File::Temp => 0},
    'dist'          => {COMPRESS => 'gzip -best', SUFFIX => 'gz'},
    'EXE_FILES'     => ['exe/chartex'],
);
