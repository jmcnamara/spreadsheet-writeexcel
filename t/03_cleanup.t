
###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Test that garbage collection of a S::WE object doesn't clober $@;
#
# See: http://groups.google.com/group/spreadsheet-writeexcel/browse_thread/thread/f5007499fc381870
#
# reverse('©'), January 2007, John McNamara, jmcnamara@cpan.org
#


use strict;

use Spreadsheet::WriteExcel;
use Test::More tests => 2;


###############################################################################
#
# Tests setup
#
my $test_file   = 'temp_test_file.xls';
my $die_message = "__SWE_test_message__\n";


###############################################################################
#
# Test Spreadsheet::WriteExcel
#
eval {
    my $workbook = Spreadsheet::WriteExcel->new($test_file);

    die $die_message;

    # $workbook goes out of scope here and will be garbage collected.
};

is ($@, $die_message, " \tCatching die message.");


###############################################################################
#
# Test Spreadsheet::WriteExcel::Big if possible.
#
SKIP: {
    eval { require OLE::Storage_Lite };
    skip "\tTesting ::Big requires OLE::Storage_Lite", 1 if $@;

    require Spreadsheet::WriteExcel::Big;

    eval {
        my $workbook = Spreadsheet::WriteExcel::Big->new($test_file);

        die $die_message;

        # $workbook goes out of scope here and will be garbage collected.
    };

    is( $@, $die_message, " \tCatching die message." );
}



unlink $test_file;


__END__
