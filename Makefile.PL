use ExtUtils::MakeMaker;

WriteMakefile(
    'AUTHOR'        => 'John McNamara (jmcnamara@cpan.org)',
    'ABSTRACT'      => 'Write to a cross platform Excel binary file',
    'NAME'          => 'Spreadsheet::WriteExcel',
    'VERSION_FROM'  => 'WriteExcel.pm',
    'NEEDS_LINKING' => 0,
    'PREREQ_PM'     => {Parse::RecDescent => 0, File::Temp => 0},
    'dist'          => {COMPRESS => 'gzip -best', SUFFIX => 'gz'},
);
