use v6;

unit module Spreadsheet::Raw::Libxlsxiowrite:ver<0.0.2>:auth<cpan:FRITH>;

use NativeCall;

constant LIB = ('xlsxio_write');

class xlsxiowriter is repr('CPointer') is export { * } # libxlsx private struct

sub xlsxiowrite_get_version(int32 $pmajor is rw, int32 $pminor is rw, int32 $pmicro is rw) is native(LIB) is export { * }
sub xlsxiowrite_get_version_string(--> Str) is native(LIB) is export { * }
sub xlsxiowrite_open(Str $filename, Str $sheetname --> xlsxiowriter) is native(LIB) is export { * }
sub xlsxiowrite_close(xlsxiowriter $handle --> int32) is native(LIB) is export { * }
sub xlsxiowrite_set_detection_rows(xlsxiowriter $handle, size_t $rows) is native(LIB) is export { * }
sub xlsxiowrite_set_row_height(xlsxiowriter $handle, size_t $height) is native(LIB) is export { * }
sub xlsxiowrite_add_column(xlsxiowriter $handle, Str $name, int32 $width) is native(LIB) is export { * }
sub xlsxiowrite_add_cell_string(xlsxiowriter $handle, Str $value) is native(LIB) is export { * }
sub xlsxiowrite_add_cell_int(xlsxiowriter $handle, int64 $value) is native(LIB) is export { * }
sub xlsxiowrite_add_cell_float(xlsxiowriter $handle, num64 $value) is native(LIB) is export { * }
sub xlsxiowrite_add_cell_datetime(xlsxiowriter $handle, int64 $value) is native(LIB) is export { * }
sub xlsxiowrite_next_row(xlsxiowriter $handle) is native(LIB) is export { * }
