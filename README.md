Super simple, super fast stream-based XLSX parsing.  Suitable for very large
files.

Requires Ruby 2.0.

    require 'dullard' 

    workbook = Dullard::Workbook.new "file.xlsx"
    workbook.sheets[0].rows.each do |row|
      puts row # => ["a","b","c",...]
    end
