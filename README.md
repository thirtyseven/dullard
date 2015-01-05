# dullard

Super simple, super fast stream-based XLSX parsing.  Suitable for very large
files.

Requires Ruby 2.0.

    require 'dullard' 

    workbook = Dullard::Workbook.new "file.xlsx"
    workbook.sheets[0].rows.each do |row|
      p row # => ["a","b","c", 0.3, #<DateTime: -4712-01-01....>, ...]
    end

## Current limitations
 * Limited validation and error handling.
 * Formatted cells are read minus formatting.
 * Rows that end with empty cells may be truncated.
 * May be buggy.  Pull requests welcome!
