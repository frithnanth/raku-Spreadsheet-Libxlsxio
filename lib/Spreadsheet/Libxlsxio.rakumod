unit class Spreadsheet::Libxlsxio:ver<0.0.3>:auth<zef:FRITH>;

use Spreadsheet::Raw::Libxlsxioread;
use Spreadsheet::Raw::Libxlsxiowrite;
use Spreadsheet::Libxlsxio::Constants;
use Spreadsheet::Libxlsxio::Exception;
use NativeCall;

class Read {
  has xlsxioreader $!rhandler;

  multi method new(Str  $file!) { self.bless(:$file) }
  multi method new(Str :$file!) { self.bless(:$file) }
  multi method new(Buf  $data!) { self.bless(:$data) }
  multi method new(Buf :$data!) { self.bless(:$data) }

  submethod BUILD(Str :$file, Buf :$data) {
    fail X::Libxlsxio.new: errno => 1, error => 'File not found' if $file.defined && ! $file.IO.f;
    with $file {
      $!rhandler = xlsxioread_open($file);
    } orwith $data {
      my $buf = nativecast(Pointer, $data);
      $!rhandler = xlsxioread_open_memory($buf, $data.bytes, 1);
    }
  }

  submethod DESTROY {
    xlsxioread_close($!rhandler) with $!rhandler;
  }

  method version(--> Str) {
    my int32 ($major, $minor, $micro);
    xlsxioread_get_version($major, $minor, $micro);
    my $version = xlsxioread_get_version_string;
    return $version but List.new: $major, $minor, $micro;
  }

  method sheet-open(Str :$name, Int :$flag = XLSXIOREAD_SKIP_EMPTY_ROWS --> xlsxioreadersheet) {
    xlsxioread_sheet_open($!rhandler, $name, $flag);
  }

  method sheet-close(xlsxioreadersheet $sheet) {
    xlsxioread_sheet_close($sheet);
  }

  method sheet-flags(xlsxioreadersheet $sheet --> UInt) {
    with $sheet {
      xlsxioread_sheet_flags($sheet) +& 127;
    } else {
      fail X::Libxlsxio.new: errno => 11, error => 'No open sheet found';
    }
  }

  method sheet-list(&callback) {
    xlsxioread_list_sheets($!rhandler, -> $name, $dummy { &callback($name) }, $!rhandler);
  }

  method sheetlist-open(--> xlsxioreadersheetlist) {
    xlsxioread_sheetlist_open($!rhandler);
  }

  method sheetlist-close(xlsxioreadersheetlist $sheetlisthandler) {
    xlsxioread_sheetlist_close($sheetlisthandler);
  }

  method sheetlist-next(xlsxioreadersheetlist $sheetlisthandler --> Str) {
    xlsxioread_sheetlist_next($sheetlisthandler);
  }

  method last-row-read(xlsxioreadersheet $sheet --> UInt) {
    xlsxioread_sheet_last_row_index($sheet);
  }

  method next-cell(xlsxioreadersheet $sheet --> Str) {
    my Pointer $value = xlsxioread_sheet_next_cell($sheet);
    my Str $strval = nativecast(Str, $value);
    xlsxioread_free($value);
    return $strval;
  }

  method last-column-read(xlsxioreadersheet $sheet --> UInt) {
    xlsxioread_sheet_last_column_index($sheet);
  }

  method next-row(xlsxioreadersheet $sheet --> UInt) {
    xlsxioread_sheet_next_row($sheet);
  }

  method process(Str $sheetname, &cell-callback, &row-callback, UInt :$flags = XLSXIOREAD_SKIP_EMPTY_ROWS --> Int) {
    xlsxioread_process($!rhandler,
                       $sheetname,
                       $flags,
                       -> $row1, $col, $value, $dummy1 { &cell-callback($row1, $col, $value) },
                       -> $row2, $maxcol, $dummy2      { &row-callback($row2, $maxcol) },
                       $!rhandler);
  }

  sub to-date(Str() $value --> DateTime) is export {
    my DateTime $d .= new('1900-01-01T00:00:00').later: :days($value - 2);
  }
}

class Write {
  has xlsxiowriter $!whandler;

  multi method new(Str  $file!, Str  $sheet?) { self.bless(:$file, :$sheet) }
  multi method new(Str :$file,  Str :$sheet)  { self.bless(:$file, :$sheet) }

  submethod BUILD(Str :$file, Str :$sheet) {
    $!whandler = xlsxiowrite_open($file, $sheet);
  }

  method close {
    xlsxiowrite_close($!whandler) with $!whandler;
  }

  method version(--> Str) {
    my int32 ($major, $minor, $micro);
    xlsxiowrite_get_version($major, $minor, $micro);
    my $version = xlsxiowrite_get_version_string;
    return $version but List.new: $major, $minor, $micro;
  }

  method detection-rows(UInt $rows --> Spreadsheet::Libxlsxio::Write) {
    xlsxiowrite_set_detection_rows($!whandler, $rows);
    self;
  }

  method row-height(UInt $height --> Spreadsheet::Libxlsxio::Write) {
    xlsxiowrite_set_row_height($!whandler, $height);
    self;
  }

  method add-column(Str $name, UInt $width --> Spreadsheet::Libxlsxio::Write) {
    xlsxiowrite_add_column($!whandler, $name, $width);
    self;
  }

  method write-row(--> Spreadsheet::Libxlsxio::Write) {
    xlsxiowrite_next_row($!whandler);
    self;
  }

  method add-string(Str() $value --> Spreadsheet::Libxlsxio::Write) {
    xlsxiowrite_add_cell_string($!whandler, $value);
    self;
  }

  method add-int(Int() $value --> Spreadsheet::Libxlsxio::Write) {
    my int64 $val = $value;
    xlsxiowrite_add_cell_int($!whandler, $val);
    self;
  }

  method add-num(Num() $value --> Spreadsheet::Libxlsxio::Write) {
    my num64 $val = $value;
    xlsxiowrite_add_cell_float($!whandler, $val);
    self;
  }

  method add-datetime(DateTime() $value --> Spreadsheet::Libxlsxio::Write) {
    my int64 $val = $value.Instant.to-posix[0].Int;
    xlsxiowrite_add_cell_datetime($!whandler, $val);
    self;
  }
}

=begin pod

=head1 NAME

Spreadsheet::Libxlsxio - An interface to libxlsxio, a C library to read and write XLSX files

=head1 SYNOPSIS

=begin code :lang<raku>

use Spreadsheet::Libxlsxio;
use Spreadsheet::Libxlsxio::Constants;

my @values;

my $ss = Spreadsheet::Libxlsxio::Read.new: 'mydata.xlsx';
$ss.sheet-list(-> $name {
  @values.push: "[$name]";
  $ss.process: $name,
               -> $, $, $value { @values.push: $value with $value; XLSXIOREAD_CONTINUE },
               -> $, $ { XLSXIOREAD_CONTINUE }
});

# Do something with @values

=end code

=head1 DESCRIPTION

Spreadsheet::Libxlsxio is an interface to libxlsxio, a C library to read and write XLSX files

libxlsxio is a fast C library released under the terms of the MIT License.
It has a nice API, which translates easily using Raku pointy blocks (or closures, or lambdas).

Please check the limitations of this library here: L<https://brechtsanders.github.io/xlsxio/>.

This module exports two classes:

=item Spreadsheet::Libxlsxio::Read
=item Spreadsheet::Libxlsxio::Write

The Read class offers two ways to read an XLSX file:

=item iterating through rows and cells;
=item using callback functions for each cell and after each row.

The Write class has methods for writing one row at the time, cell after cell, with no random access to a single cell.

=head2 Spreadsheet::Libxlsxio::Read

=head3 new(Str  $file!)
=head3 new(Str :$file!)
=head3 new(Buf  $data!)
=head3 new(Buf :$data!)

This B<multi method> constructor requires one simple or named argument, the file name or the B<Buf> variable into which the spreadsheet data has been read.

=head3 version(--> Str)

This method returns the C library version as a B<Str> or a B<List>, if evaluated in a List context.

=begin code

my $ss = Spreadsheet::Libxlsxio::Read.new: :file('mydata.xlsx');
my $version = $ss.version;
say $version;      # '0.2.29'
say $version.List; # (0 2 29)

=end code

B<Methods to iterate through worksheets, rows, and cells>

=head3 sheetlist-open(--> xlsxioreadersheetlist)

This method opens the worksheet list.
It return the worksheet list handler.

=head3 sheetlist-next(xlsxioreadersheetlist $sheetlisthandler --> Str)

This method accepts a worksheet list handler and returns the name of the next available worksheet.

For example:

=begin code
my $ss = Spreadsheet::Libxlsxio::Read.new: :file('mydata.xlsx');
my $sl = $ss.sheetlist-open;
while $ss.sheetlist-next($sl) -> $sheetname { say $sheetname }
$ss.sheetlist-close($sl);
=end code

=head3 sheetlist-close(xlsxioreadersheetlist $sheetlisthandler)

Closes the worksheet list handler.

=head3 sheet-open(Str :$name, Int :$flag = XLSXIOREAD_SKIP_EMPTY_ROWS --> xlsxioreadersheet)

Opens a worksheet by name. The flag values are defined in the B<Spreadsheet::Libxlsxio::Constants> module:

=item XLSXIOREAD_SKIP_NONE
=item XLSXIOREAD_SKIP_EMPTY_ROWS
=item XLSXIOREAD_SKIP_EMPTY_CELLS
=item XLSXIOREAD_SKIP_ALL_EMPTY
=item XLSXIOREAD_SKIP_EXTRA_CELLS

This method returns a handler to the specified worksheet.

=head3 sheet-close(xlsxioreadersheet $sheet)

This method takes the worksheet handler and closes the worksheet.

=head3 sheet-flags(xlsxioreadersheet $sheet --> UInt)

This method takes the worksheet handler and returns the flags set when opening the worksheet.

=head3 last-row-read(xlsxioreadersheet $sheet --> UInt)

This method takes the worksheet handler and returns the index of the last row read.

=head3 last-column-read(xlsxioreadersheet $sheet --> UInt)

This method takes the worksheet handler and returns the index of the last column read.

=head3 next-cell(xlsxioreadersheet $sheet --> Str)

This method takes the worksheet handler and returns the value of the last cell read, or a false value if there are no more cells available in the current row.

=head3 next-row(xlsxioreadersheet $sheet --> UInt)

This method takes the worksheet handler and gets the next row, or a false value if there are no more rows available.

This method must be called before each row.


B<Methods that use callback functions for each cell and after each row>

=head3 sheet-list(&callback)

This method sets a callback to be called on each worksheet.

The callback function receives the worksheet name. The actual processing can be done by the following method.

=head3 process(Str $sheetname, &cell-callback, &row-callback, UInt :$flags = XLSXIOREAD_SKIP_EMPTY_ROWS --> Int)

This method processes all rows and columns of a worksheet in a XLSX file.

It takes the sheet name, two callback functions and a flag and returns zero on success or non-zero on error.
The first callback is called when a cell value is available; it receives three arguments: the row, the column, and the value contained in the cell. It must return one of these two values, available defined in the B<Spreadsheet::Libxlsxio::Constants> module:

=item XLSXIOREAD_CONTINUE
=item XLSXIOREAD_ABORT

The second callback is called after each row; it receives two arguments: the row number (starting from 1) and the maximum column number on the row (starting from 1).

For example one might use both the previous methods to scan the entire spreadsheet:

=begin code
my $ss = Spreadsheet::Libxlsxio::Read.new: 't/test.xlsx';

my @values;

$ss.sheet-list(-> $name {
  note "Processing worksheet [$name]";
  $ss.process: $name,
               -> $, $, $value { @values.push: $value with $value; XLSXIOREAD_CONTINUE },
               -> $, $ { XLSXIOREAD_CONTINUE }
});

dd @values;

=end code

B<NOTE>

All the cell values are returned as B<Str>.
A helper sub is provided to convert a date cell to a Raku DateTime object:

=head3 sub to-date(Str() $value --> DateTime)

=head2 Spreadsheet::Libxlsxio::Write

=begin code

use Spreadsheet::Libxlsxio;

my $ss = Spreadsheet::Libxlsxio::Write.new: :file('mydata.xlsx'), :sheet('Sheet1');
$ss.detection-rows(10)
   .row-height(10)
   .add-column('Col1', 16)
   .add-column('Col2', 0)
   .add-column('Col3', 0)
   .add-column('Col4', 0)
   .write-row;
my $dt = DateTime.new('2021-07-25T00:00:00').Instant.to-posix[0];
for ^10 -> $i {
  $ss.add-string("Test$i").add-int($i).add-num(π).add-datetime($dt).write-row;
}
$ss.close;

=end code

=head3 new(Str  $file!, Str  $sheet?)
=head3 new(Str :$file,  Str :$sheet)

This B<multi method> constructor requires two simple or named arguments: the file name and the worksheet name.

=head3 close

This method closes the spreadsheet file.

=head3 version(--> Str)

This method returns the C library version as a B<Str> or a B<List>, if evaluated in a List context.

=begin code

my $ss = Spreadsheet::Libxlsxio::Write.new: :file('mydata.xlsx');
my $version = $ss.version;
say $version;      # '0.2.xx'
say $version.List; # (0 2 xx)

=end code

B<All the following methods return their object, so they may be chained.>

=head3 detection-rows(UInt $rows --> Spreadsheet::Libxlsxio::Write)

This method specifies how many initial rows will be buffered in memory to determine column widths.

=head3 row-height(UInt $height --> Spreadsheet::Libxlsxio::Write)

This method specifies the row height in text lines to use for the current and next rows.

=head3 add-column(Str $name, UInt $width --> Spreadsheet::Libxlsxio::Write)

This method adds a column label cell. It takes two arguments: the label and the column width in characters.
It must be called for each column; the row is committed by calling the B<write-row> method.

=head3 write-row(--> Spreadsheet::Libxlsxio::Write)

This method marks the end of a row.

=head3 add-string(Str() $value --> Spreadsheet::Libxlsxio::Write)

This method adds a cell containing Str data.

=head3 add-int(Int() $value --> Spreadsheet::Libxlsxio::Write)

This method adds a cell containing Int data.

=head3 add-num(Num() $value --> Spreadsheet::Libxlsxio::Write)

This method adds a cell containing Num data.

=head3 add-datetime(DateTime() $value --> Spreadsheet::Libxlsxio::Write)

This method adds a cell containing DateTime data.

=head1 Installation

To use this module one has to install the libxlsxio C library first.

The library home page is here: L<https://brechtsanders.github.io/xlsxio/> and the GitHub project page is
L<https://github.com/brechtsanders/xlsxio>.

The C library has two dependencies: B<expat> (available as libexpat1 on Ubuntu and Debian Linux systems) and one of B<minizip> or B<libzip> (libminizip1 and libzip5 on Ubuntu or Debian Linux systems).

Installing the library is as simple as typing:

=begin input
sudo make install
sudo ldconfig
=end input

To install the Raku module using zef (a module management tool):

=begin code
$ zef install Spreadsheet::Libxlsxio
=end code

=head1 AUTHOR

Fernando Santagata <nando.santagata@gmail.com>

=head1 COPYRIGHT AND LICENSE

Copyright 2021 Fernando Santagata

This library is free software; you can redistribute it and/or modify it under the Artistic License 2.0.

=end pod
