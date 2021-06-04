use v6;

unit module Spreadsheet::Raw::Libxlsxioread::Constants:ver<0.0.1>:auth<cpan:FRITH>;

constant XLSXIOREAD_SKIP_NONE         is export = 0x00;
constant XLSXIOREAD_SKIP_EMPTY_ROWS   is export = 0x01;
constant XLSXIOREAD_SKIP_EMPTY_CELLS  is export = 0x02;
constant XLSXIOREAD_SKIP_ALL_EMPTY    is export = XLSXIOREAD_SKIP_EMPTY_ROWS +| XLSXIOREAD_SKIP_EMPTY_CELLS;
constant XLSXIOREAD_SKIP_EXTRA_CELLS  is export = 0x04;

constant XLSXIOREAD_CONTINUE          is export = 0;
constant XLSXIOREAD_ABORT             is export = 1;
