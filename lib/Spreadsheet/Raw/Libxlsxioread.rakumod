use v6;

unit module Spreadsheet::Raw::Libxlsxioread:ver<0.0.2>:auth<cpan:FRITH>;

use NativeCall;

constant LIB = ('xlsxio_read');

# libxlsx private structures
class xlsxioreader is repr('CPointer') is export { * }
class xlsxioreadersheetlist is repr('CPointer') is export { * }
class xlsxioreadersheet is repr('CPointer') is export { * }

sub xlsxioread_get_version(int32 $pmajor is rw, int32 $pminor is rw, int32 $pmicro is rw) is native(LIB) is export { * }
sub xlsxioread_get_version_string(--> Str) is native(LIB) is export { * }
sub xlsxioread_open(Str $filename --> xlsxioreader) is native(LIB) is export { * }
sub xlsxioread_open_memory(Pointer $data, uint64 $datalen, int32 $freedata --> xlsxioreader) is native(LIB) is export { * }
sub xlsxioread_close(xlsxioreader $handle) is native(LIB) is export { * }
sub xlsxioread_list_sheets(xlsxioreader $handle,
                            &callback (Str $name, Pointer $callbackdata1 --> int32),
                            Pointer $callbackdata2) is native(LIB) is export { * }
sub xlsxioread_process(xlsxioreader $handle, Str $sheetname, uint32 $flags,
                            &cell-callback (size_t $row1, size_t $col, Str $value, Pointer $callbackdata1 --> int32),
                            &row-callback (size_t $row2, size_t $maxcol, Pointer $callbackdata2),
                            Pointer $callbackdata --> int32) is native(LIB) is export { * }
sub xlsxioread_sheetlist_open(xlsxioreader $handle, --> xlsxioreadersheetlist) is native(LIB) is export { * }
sub xlsxioread_sheetlist_close(xlsxioreadersheetlist $sheetlisthandle) is native(LIB) is export { * }
sub xlsxioread_sheetlist_next(xlsxioreadersheetlist $sheetlisthandle --> Str) is native(LIB) is export { * }
sub xlsxioread_sheet_last_row_index(xlsxioreadersheet $sheethandle --> size_t) is native(LIB) is export { * }
sub xlsxioread_sheet_last_column_index(xlsxioreadersheet $sheethandle --> size_t) is native(LIB) is export { * }
sub xlsxioread_sheet_flags(xlsxioreadersheet $sheethandle --> uint32) is native(LIB) is export { * }
sub xlsxioread_sheet_open(xlsxioreader $handle, Str $sheetname, uint32 $flags --> xlsxioreadersheet) is native(LIB) is export { * }
sub xlsxioread_sheet_close(xlsxioreadersheet $sheethandle) is native(LIB) is export { * }
sub xlsxioread_sheet_next_row(xlsxioreadersheet $sheethandle --> int32) is native(LIB) is export { * }
sub xlsxioread_sheet_next_cell(xlsxioreadersheet $sheethandle --> Pointer) is native(LIB) is export { * }
sub xlsxioread_free(Pointer $data) is native(LIB) is export { * }
