#!/usr/bin/env raku

# This is the translation of the C library example_xlsxio_read_advanced.c

use Spreadsheet::Libxlsxio;
use Spreadsheet::Libxlsxio::Constants;

unit sub MAIN(Str $filename where *.IO.f = 'example.xlsx');

# open .xlsx file for reading
my $ss = Spreadsheet::Libxlsxio::Read.new: $filename;

# list available sheets
$ss.sheet-list( -> $sheetname {
  say " - $sheetname";
  $ss.process: $sheetname,
               -> $, $col, $value {
                                    print "\t" if $col > 1;
                                    print $value with $value;
                                    XLSXIOREAD_CONTINUE
                                  },
               -> $, $ {
                          print "\n";
                          XLSXIOREAD_CONTINUE
                       }
                  }
              );
