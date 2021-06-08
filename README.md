[![Actions Status](https://github.com/frithnanth/raku-Spreadsheet-Libxlsxio/workflows/test/badge.svg)](https://github.com/frithnanth/raku-Spreadsheet-Libxlsxio/actions)

NAME
====

Spreadsheet::Libxlsxio - An interface to libxlsxio, a C library to read and write XLSX files

SYNOPSIS
========

```raku
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
```

DESCRIPTION
===========

Spreadsheet::Libxlsxio is an interface to libxlsxio, a C library to read and write XLSX files

libxlsxio is a fast C library released under the terms of the MIT License. It has a nice API, which translates easily using Raku pointy blocks (or closures, or lambdas).

Please check the limitations of this library here: [https://brechtsanders.github.io/xlsxio/](https://brechtsanders.github.io/xlsxio/).

This module exports two classes:

  * Spreadsheet::Libxlsxio::Read

  * Spreadsheet::Libxlsxio::Write

The Read class offers two ways to read an XLSX file:

  * iterating through rows and cells;

  * using callback functions for each cell and after each row.

The Write class has methods for writing one row at the time, cell after cell, with no random access to a single cell.

Spreadsheet::Libxlsxio::Read
----------------------------

### new(Str $file!)

### new(Str :$file!)

### new(Buf $data!)

### new(Buf :$data!)

This **multi method** constructor requires one simple or named argument, the file name or the **Buf** variable into which the spreadsheet data has been read.

### version(--> Str)

This method returns the C library version as a **Str** or a **List**, if evaluated in a List context.

    my $ss = Spreadsheet::Libxlsxio::Read.new: :file('mydata.xlsx');
    my $version = $ss.version;
    say $version;      # '0.2.29'
    say $version.List; # (0 2 29)

**Methods to iterate through worksheets, rows, and cells**

### sheetlist-open(--> xlsxioreadersheetlist)

This method opens the worksheet list. It return the worksheet list handler.

### sheetlist-next(xlsxioreadersheetlist $sheetlisthandler --> Str)

This method accepts a worksheet list handler and returns the name of the next available worksheet.

For example:

    my $ss = Spreadsheet::Libxlsxio::Read.new: :file('mydata.xlsx');
    my $sl = $ss.sheetlist-open;
    while $ss.sheetlist-next($sl) -> $sheetname { say $sheetname }
    $ss.sheetlist-close($sl);

### sheetlist-close(xlsxioreadersheetlist $sheetlisthandler)

Closes the worksheet list handler.

### sheet-open(Str :$name, Int :$flag = XLSXIOREAD_SKIP_EMPTY_ROWS --> xlsxioreadersheet)

Opens a worksheet by name. The flag values are defined in the **Spreadsheet::Libxlsxio::Constants** module:

  * XLSXIOREAD_SKIP_NONE

  * XLSXIOREAD_SKIP_EMPTY_ROWS

  * XLSXIOREAD_SKIP_EMPTY_CELLS

  * XLSXIOREAD_SKIP_ALL_EMPTY

  * XLSXIOREAD_SKIP_EXTRA_CELLS

This method returns a handler to the specified worksheet.

### sheet-close(xlsxioreadersheet $sheet)

This method takes the worksheet handler and closes the worksheet.

### sheet-flags(xlsxioreadersheet $sheet --> UInt)

This method takes the worksheet handler and returns the flags set when opening the worksheet.

### last-row-read(xlsxioreadersheet $sheet --> UInt)

This method takes the worksheet handler and returns the index of the last row read.

### last-column-read(xlsxioreadersheet $sheet --> UInt)

This method takes the worksheet handler and returns the index of the last column read.

### next-cell(xlsxioreadersheet $sheet --> Str)

This method takes the worksheet handler and returns the value of the last cell read, or a false value if there are no more cells available in the current row.

### next-row(xlsxioreadersheet $sheet --> UInt)

This method takes the worksheet handler and gets the next row, or a false value if there are no more rows available.

This method must be called before each row.

**Methods that use callback functions for each cell and after each row**

### sheet-list(&callback)

This method sets a callback to be called on each worksheet.

The callback function receives the worksheet name. The actual processing can be done by the following method.

### process(Str $sheetname, &cell-callback, &row-callback, UInt :$flags = XLSXIOREAD_SKIP_EMPTY_ROWS --> Int)

This method processes all rows and columns of a worksheet in a XLSX file.

It takes the sheet name, two callback functions and a flag and returns zero on success or non-zero on error. The first callback is called when a cell value is available; it receives three arguments: the row, the column, and the value contained in the cell. It must return one of these two values, available defined in the **Spreadsheet::Libxlsxio::Constants** module:

  * XLSXIOREAD_CONTINUE

  * XLSXIOREAD_ABORT

The second callback is called after each row; it receives two arguments: the row number (starting from 1) and the maximum column number on the row (starting from 1).

For example one might use both the previous methods to scan the entire spreadsheet:

    my $ss = Spreadsheet::Libxlsxio::Read.new: 't/test.xlsx';

    my @values;

    $ss.sheet-list(-> $name {
      note "Processing worksheet [$name]";
      $ss.process: $name,
                   -> $, $, $value { @values.push: $value with $value; XLSXIOREAD_CONTINUE },
                   -> $, $ { XLSXIOREAD_CONTINUE }
    });

    dd @values;

**NOTE**

All the cell values are returned as **Str**. A helper sub is provided to convert a date cell to a Raku DateTime object:

### sub to-date(Str() $value --> DateTime)

Spreadsheet::Libxlsxio::Write
-----------------------------

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

### new(Str $file!, Str $sheet?)

### new(Str :$file, Str :$sheet)

This **multi method** constructor requires two simple or named arguments: the file name and the worksheet name.

### close

This method closes the spreadsheet file.

### version(--> Str)

This method returns the C library version as a **Str** or a **List**, if evaluated in a List context.

    my $ss = Spreadsheet::Libxlsxio::Write.new: :file('mydata.xlsx');
    my $version = $ss.version;
    say $version;      # '0.2.29'
    say $version.List; # (0 2 29)

**All the following methods return their object, so they may be chained.**

### detection-rows(UInt $rows --> Spreadsheet::Libxlsxio::Write)

This method specifies how many initial rows will be buffered in memory to determine column widths.

### row-height(UInt $height --> Spreadsheet::Libxlsxio::Write)

This method specifies the row height in text lines to use for the current and next rows.

### add-column(Str $name, UInt $width --> Spreadsheet::Libxlsxio::Write)

This method adds a column label cell. It takes two arguments: the label and the column width in characters. It must be called for each column; the row is committed by calling the **write-row** method.

### write-row(--> Spreadsheet::Libxlsxio::Write)

This method marks the end of a row.

### add-string(Str() $value --> Spreadsheet::Libxlsxio::Write)

This method adds a cell containing Str data.

### add-int(Int() $value --> Spreadsheet::Libxlsxio::Write)

This method adds a cell containing Int data.

### add-num(Num() $value --> Spreadsheet::Libxlsxio::Write)

This method adds a cell containing Num data.

### add-datetime(DateTime() $value --> Spreadsheet::Libxlsxio::Write)

This method adds a cell containing DateTime data.

Installation
============

To use this module one has to install the libxlsxio C library first.

The library home page is here: [https://brechtsanders.github.io/xlsxio/](https://brechtsanders.github.io/xlsxio/) and the GitHub project page is [https://github.com/brechtsanders/xlsxio](https://github.com/brechtsanders/xlsxio).

The C library has two dependencies: **expat** (available as libexpat1 on Ubuntu and Debian Linux systems) and one of **minizip** or **libzip** (libminizip1 and libzip5 on Ubuntu or Debian Linux systems).

Installing the library is as simple as typing:

    sudo make install
    sudo ldconfig

To install the Raku module using zef (a module management tool):

    $ zef install Spreadsheet::Libxlsxio

AUTHOR
======

Fernando Santagata <nando.santagata@gmail.com>

COPYRIGHT AND LICENSE
=====================

Copyright 2021 Fernando Santagata

This library is free software; you can redistribute it and/or modify it under the Artistic License 2.0.

