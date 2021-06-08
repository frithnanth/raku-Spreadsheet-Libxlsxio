#!/usr/bin/env raku

# This is the translation of the C library example_xlsxio_read.c

use Spreadsheet::Libxlsxio;

unit sub MAIN(Str $filename where *.IO.f = 'example.xlsx');

# open .xlsx file for reading
my $ss = Spreadsheet::Libxlsxio::Read.new: $filename;
note "XLSX I/O library version { $ss.version }";

# list available sheets
say "Available sheets:";
my $sl = $ss.sheetlist-open;
while $ss.sheetlist-next($sl) -> $sheetname { say " - $sheetname" }
$ss.sheetlist-close($sl);

# read values from first sheet
say "Contents of first sheet:";
my $sheet = $ss.sheet-open;

while $ss.next-row($sheet) {
  while $ss.next-cell($sheet) -> $value {
    print "$value\t";
  }
  print "\n";
}

$ss.sheet-close($sheet);
