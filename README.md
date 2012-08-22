Super simple, super fast XLSX parsing.

    require 'dullard' 

    workbook = Dullard::Workbook.new "file"
    workbook.sheets[0].rows.each do |row|
      puts row # => ["a","b","c",...]
    end
