#!/usr/bin/env raku

# This is the translation of the C library example_xlsxio_write.c

use Spreadsheet::Libxlsxio;

unit sub MAIN(Str $filename where ! *.IO.f = 'example.xlsx');

my $ss = Spreadsheet::Libxlsxio::Write.new: :file($filename), :sheet('MySheet');
$ss.row-height(1)
   .detection-rows(10)
   .add-column("Col1", 0)
   .add-column("Col2", 21)
   .add-column("Col3", 0)
   .add-column("Col4", 2)
   .add-column("Col5", 0)
   .add-column("Col6", 0)
   .add-column("Col7", 0)
   .write-row;
for ^1_000 -> $i {
  $ss.add-string('Test')
     .add-string("A b  c   d    e     f\nnew line")
     .add-string("&% <test> \"'")
     .add-string(Str)
     .add-int($i)
     .add-datetime(DateTime.now)
     .add-num(π)
     .write-row;
}
$ss.close;
