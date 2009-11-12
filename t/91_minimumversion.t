#!/usr/bin/perl

# Test that our declared minimum Perl version matches our syntax
use strict;

BEGIN {
	$|  = 1;
	$^W = 1;
}

my @MODULES = (
	'Perl::MinimumVersion 1.20',
	'Test::MinimumVersion 0.008',
);

# Don't run tests during end-user installs
use Test::More;
unless ( $ENV{AUTOMATED_TESTING} or $ENV{RELEASE_TESTING} ) {
	plan( skip_all => "Author tests not required for installation" );
}

# Load the testing modules
foreach my $MODULE ( @MODULES ) {
	eval "use $MODULE";
	if ( $@ ) {
		$ENV{RELEASE_TESTING}
		? die( "Failed to load required release-testing module $MODULE" )
		: plan( skip_all => "$MODULE not available for testing" );
	}
}


skip_some_minimum_version_ok('5.005');

###############################################################################
#
# This is a modified version of minimum_version_ok() from Test::MinimumVersion
# that skips modules that return false minumum versions due to the following
# issue in Perl::MinimumVersion:
#     http://rt.cpan.org/Public/Bug/Display.html?id=51256
#
sub skip_some_minimum_version_ok {
  my ($version, $arg) = @_;
  $arg ||= {};
  $arg->{paths} ||= [ qw(lib t xt/smoke), glob ("*.pm"), glob ("*.PL") ];

  my $Test = Test::Builder->new;

  $version = Test::MinimumVersion::_objectify_version($version);

  my @perl_files;
  for my $path (@{ $arg->{paths} }) {
    if (-f $path and -s $path) {
      push @perl_files, $path;
    } elsif (-d $path) {
      push @perl_files, File::Find::Rule->perl_file->in($path);
    }
  }

  unless ($Test->has_plan or $arg->{no_plan}) {
    $Test->plan(tests => scalar @perl_files);
  }

  for (@perl_files) {

      SKIP: {
          skip 'due to RT 51256 issue in Perl::Minimum version', 1
          if /(OLEwriter.pm|Workbook.pm)$/;

          minimum_version_ok($_, $version);
      }
  }
}


1;
