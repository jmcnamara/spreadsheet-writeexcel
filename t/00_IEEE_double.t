
#####################################################
# We start with some black magic to print on failure.

BEGIN { $| = 1; print "1..2\n"; }
END {print "not ok 1\n" unless $loaded;}
use Spreadsheet::WriteExcel;
$loaded = 1;
print "ok 1\n";

#####################################################
# End of black magic.



# TEST 2
#
# Check if "pack" gives the required IEEE 64bit float
my $teststr = pack "d", 1.2345;
my @hexdata = (0x8D, 0x97, 0x6E, 0x12, 0x83, 0xC0, 0xF3, 0x3F);
my $number  = pack "C8", @hexdata;

if ($number eq $teststr) {
    # Little Endian
    print "ok 2\n";
}
elsif ($number eq reverse($teststr)){
    # Big Endian
    print "ok 2\n";
}
else {
    # Give up. I'll fix this in a later version.
    print "not ok 2\n";
}
