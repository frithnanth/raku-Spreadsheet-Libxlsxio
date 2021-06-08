#!/usr/bin/env raku

# This is the translation of the C library example_xlsxio_write_getversion.c

use Spreadsheet::Libxlsxio;

unit sub MAIN;

my $ss = Spreadsheet::Libxlsxio::Write.new(Str).version;
say 'Version: ' ~ $ss.List.join('.');
say "Version: $ss";
say "libxlsxio_write $ss";
