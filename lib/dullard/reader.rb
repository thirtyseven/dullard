require 'zip/zipfilesystem'
require 'nokogiri'

module Dullard; end

class Dullard::Workbook
  def initialize(file)
    @file = file
    @zipfs = Zip::ZipFile.open(@file)
  end

  def sheets
    workbook = Nokogiri::XML::Document.parse(@zipfs.file.open("xl/workbook.xml"))
    @sheets = workbook.css("sheet").each_with_index.map {|n,i| Dullard::Sheet.new(self, n.attr("name"), n.attr("sheetId"), i+1) }
  end

  def string_table
    @string_tabe ||= read_string_table
  end

  def read_string_table
    @string_table = []
    entry = ''
    Nokogiri::XML::Reader(@zipfs.file.open("xl/sharedStrings.xml")).each do |node|
      if node.name == "si" and node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT
        entry = ''
      elsif node.name == "si" and node.node_type == Nokogiri::XML::Reader::TYPE_END_ELEMENT
        @string_table << entry
      elsif node.value?
        entry << node.value
      end
    end
    @string_table
  end

  def zipfs
    @zipfs
  end

  def close
    @zipfs.close
  end
end

class Dullard::Sheet
  attr_reader :name, :workbook
  def initialize(workbook, name, id, index)
    @workbook = workbook
    @name = name
    @id = id
    @index = index
    @file = @workbook.zipfs.file.open(path) if @workbook.zipfs.file.exist?(path)
  end

  def string_lookup(i)
    @workbook.string_table[i]
  end

  def rows
    Enumerator.new(rows_size) do |y|
      next unless @file
      @file.rewind
      shared = false
      row = nil
      column = nil
      Nokogiri::XML::Reader(@file).each do |node|
        case node.node_type
        when Nokogiri::XML::Reader::TYPE_ELEMENT
          case node.name
          when "row"
            row = []
            column = 0
            next
          when "c"
            if rcolumn = node.attributes["r"]
              rcolumn.delete!("0-9")
              while column < self.class.column_names.size and rcolumn != self.class.column_names[column]
                row << nil
                column += 1
              end
            end
            shared = (node.attribute("t") == "s")
            column += 1
            next
          end
        when Nokogiri::XML::Reader::TYPE_END_ELEMENT
          if node.name == "row"
            y << row
            next
          end
        end
        if value = node.value
          row << (shared ? string_lookup(value.to_i) : value)
        end
      end
    end
  end

  # Returns A to ZZZ.
  def self.column_names
    if @column_names
      @column_names
    else
      proc = Proc.new do |l|
        ("#{l}A".."#{l}Z").to_a
      end
      x = proc.call(nil)
      y = x.map(&proc).flatten
      z = y.map(&proc).flatten
      @column_names = x + y + z
    end
  end

  private
  def path
    "xl/worksheets/sheet#{@index}.xml"
  end

  def rows_size
    if defined? @rows_size
      @rows_size
    elsif @file
      @file.rewind
      Nokogiri::XML::Reader(@file).each do |node|
        if node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT
          case node.name
          when "dimension"
            if ref = node.attributes["ref"]
              break @rows_size = ref.scan(/\d+$/).first.to_i
            end
          when "sheetData"
            break @rows_size = nil
          end
        end
      end
    end
  end
end
