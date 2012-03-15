require 'zip/zipfilesystem'
require 'nokogiri'

module Dullard; end

class Dullard::Workbook
  def initialize(file)
    @file = file
    @zipfs = Zip::ZipFile.open(@file)
  end

  def sheets
    @workbook = Nokogiri::XML::Document.parse(@zipfs.file.open("xl/workbook.xml"))
    @sheets = @workbook.css("sheet").map {|n| Dullard::Sheet.new(self, n.attr("name"), n.attr("sheetId")) }
  end

  def string_table
    @string_tabe ||= read_string_table
  end

  def read_string_table
    @string_table = []
    state = 0
    Nokogiri::XML::Reader(@zipfs.file.open("xl/sharedStrings.xml")).each do |node|
      case state
      when 0
        if node.name == "t"
          state = 1
        end
      when 1
        @string_table << node.value
        state = 0
      end
    end
    @string_table
  end

  def zipfs
    @zipfs
  end
end

class Dullard::Sheet
  def initialize(workbook, name, id)
    @workbook = workbook
    @name = name
    @id = id
  end

  def string_lookup(i)
    @workbook.string_table[i]
  end

  def rows
    Enumerator.new do |y|
      state = :top
      row = nil
      Nokogiri::XML::Reader(@workbook.zipfs.file.open("xl/worksheets/sheet#{@id}.xml")).each do |node|
        case state
        when :top
          if node.name == "row"
            state = :row 
            y << row unless row.nil?
            row = []
          end
        when :row
          if node.name == "row"
            y << row
            row = []
          elsif node.attribute("t") == "s"
            state = :cell_shared
          else
            state = :cell
          end
        when :cell_shared
          row << string_lookup(node.value.to_i)
          state = :row
        when :cell
          row << node.value
          state = :row
        end
      end
    end
  end
end

