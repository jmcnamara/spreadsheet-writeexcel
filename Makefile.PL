use ExtUtils::MakeMaker;

    @if5005 = (
        'AUTHOR'    => 'John McNamara (jmcnamara@cpan.org)',
        'ABSTRACT'  => 'Write cross platform Excel binary file',
    ) if $] ge '5.005';

    WriteMakefile(
        'NAME'          => 'Spreadsheet::WriteExcel',
        'VERSION_FROM'  => 'WriteExcel.pm',
        'NEEDS_LINKING' => 0,
        'dist' => {COMPRESS => 'gzip -best', SUFFIX => 'gz'},
     @if5005
    );
