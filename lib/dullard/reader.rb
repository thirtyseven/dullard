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
  end

  def string_lookup(i)
    @workbook.string_table[i]
  end

  def rows
    Enumerator.new do |y|
      shared = false
      row = nil
      Nokogiri::XML::Reader(@workbook.zipfs.file.open(path)).each do |node|
        if node.name == "row" and node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT
          row = []
        elsif node.name == "row" and node.node_type == Nokogiri::XML::Reader::TYPE_END_ELEMENT
          y << row
        elsif node.name == "c" and node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT
            shared = (node.attribute("t") == "s")
        elsif node.value?
            row << (shared ? string_lookup(node.value.to_i) : node.value)
        end
      end if @workbook.zipfs.file.exist?(path)
    end
  end

  private
  def path
    "xl/worksheets/sheet#{@index}.xml"
  end
end

