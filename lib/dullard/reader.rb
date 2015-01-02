require 'zip/filesystem'
require 'nokogiri'

module Dullard
  class Error < StandardError; end
end

class Dullard::Workbook
  # Code borrowed from Roo (https://github.com/hmcgowan/roo/blob/master/lib/roo/excelx.rb)
  # Some additional formats added by Paul Hendryx (phendryx@gmail.com) that are common in LibreOffice.
  FORMATS = {
    'general' => :float,
    '0' => :float,
    '0.00' => :float,
    '#,##0' => :float,
    '#,##0.00' => :float,
    '0%' => :percentage,
    '0.00%' => :percentage,
    '0.00E+00' => :float,
    '# ?/?' => :float, #??? TODO:
    '# ??/??' => :float, #??? TODO:
    'mm-dd-yy' => :date,
    'd-mmm-yy' => :date,
    'd-mmm' => :date,
    'mmm-yy' => :date,
    'h:mm am/pm' => :date,
    'h:mm:ss am/pm' => :date,
    'h:mm' => :time,
    'h:mm:ss' => :time,
    'm/d/yy h:mm' => :date,
    '#,##0 ;(#,##0)' => :float,
    '#,##0 ;[red](#,##0)' => :float,
    '#,##0.00;(#,##0.00)' => :float,
    '#,##0.00;[red](#,##0.00)' => :float,
    'mm:ss' => :time,
    '[h]:mm:ss' => :time,
    'mmss.0' => :time,
    '##0.0e+0' => :float,
    '@' => :float,
    #-- zusaetzliche Formate, die nicht standardmaessig definiert sind:
    "yyyy\\-mm\\-dd" => :date,
    'dd/mm/yy' => :date,
    'hh:mm:ss' => :time,
    "dd/mm/yy\\ hh:mm" => :datetime,
    'm/d/yy' => :date,
    'mm/dd/yy' => :date,
    'mm/dd/yyyy' => :date,
  }

  STANDARD_FORMATS = { 
    0 => 'General',
    1 => '0',
    2 => '0.00',
    3 => '#,##0',
    4 => '#,##0.00',
    9 => '0%',
    10 => '0.00%',
    11 => '0.00E+00',
    12 => '# ?/?',
    13 => '# ??/??',
    14 => 'mm-dd-yy',
    15 => 'd-mmm-yy',
    16 => 'd-mmm',
    17 => 'mmm-yy',
    18 => 'h:mm AM/PM',
    19 => 'h:mm:ss AM/PM',
    20 => 'h:mm',
    21 => 'h:mm:ss',
    22 => 'm/d/yy h:mm',
    37 => '#,##0 ;(#,##0)',
    38 => '#,##0 ;[Red](#,##0)',
    39 => '#,##0.00;(#,##0.00)',
    40 => '#,##0.00;[Red](#,##0.00)',
    45 => 'mm:ss',
    46 => '[h]:mm:ss',
    47 => 'mmss.0',
    48 => '##0.0E+0',
    49 => '@',
  }

  def initialize(file, user_defined_formats = {})
    @file = file
    begin
      @zipfs = Zip::File.open(@file)
    rescue Zip::Error => e
      raise Dullard::Error, e.message
    end
    @user_defined_formats = user_defined_formats
    read_styles
  end

  def sheets
    workbook = Nokogiri::XML::Document.parse(@zipfs.file.open('xl/workbook.xml'))
    @sheets = workbook.css('sheet').each_with_index.map do |n, i|
      Dullard::Sheet.new(self, n.attr('name'), n.attr('sheetId'), i+1)
    end
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

  def read_styles
    doc = Nokogiri::XML(@zipfs.file.open("xl/styles.xml"))
    
    @num_formats = {}
    @cell_xfs = []
    
    doc.css('/styleSheet/numFmts/numFmt').each do |numFmt|
      if numFmt.attributes['numFmtId'] && numFmt.attributes['formatCode']
        numFmtId = numFmt.attributes['numFmtId'].value.to_i
        formatCode = numFmt.attributes['formatCode'].value
        @num_formats[numFmtId] = formatCode
      end
    end

    doc.css('/styleSheet/cellXfs/xf').each do |xf|
      if xf.attributes['numFmtId']
        numFmtId = xf.attributes['numFmtId'].value.to_i
        @cell_xfs << numFmtId
      end
    end
  end

  
  # Code borrowed from Roo (https://github.com/hmcgowan/roo/blob/master/lib/roo/excelx.rb)
  # convert internal excelx attribute to a format
  def attribute2format(s)
    id = @cell_xfs[s.to_i].to_i
    result = @num_formats[id]

    if result == nil
      if STANDARD_FORMATS.has_key? id
        result = STANDARD_FORMATS[id]
      end
    end

    result.downcase
  end

  # Code borrowed from Roo (https://github.com/hmcgowan/roo/blob/master/lib/roo/excelx.rb)
  def format2type(format)
    if FORMATS.has_key? format
      FORMATS[format]
    elsif @user_defined_formats.has_key? format
      @user_defined_formats[format]
    else
      :float
    end
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
    begin
      @file = @workbook.zipfs.file.open(path) if @workbook.zipfs.file.exist?(path)
    rescue Zip::Error => e
      raise Dullard::Error, "Couldn't open sheet #{index}: #{e.message}"
    end
  end

  def string_lookup(i)
    @workbook.string_table[i]
  end

  def rows
    Enumerator.new(row_count) do |y|
      next unless @file
      @file.rewind
      shared = false
      row = nil
      cell_map = nil # Map of column letter to cell value for a row
      column = nil
      cell_type = nil
      Nokogiri::XML::Reader(@file).each do |node|
        case node.node_type
        when Nokogiri::XML::Reader::TYPE_ELEMENT
          case node.name
          when "row"
            cell_map = {}
            next
          when 'c'
            node_type = node.attributes['t']
            cell_index = node.attributes['r']
            if !cell_index
              raise Dullard::Error, 'Invalid spreadsheet XML.'
            end

            if node_type != 's' && node_type != 'b'
              cell_format_index = node.attributes['s'].to_i
              cell_type = @workbook.format2type(@workbook.attribute2format(cell_format_index))
            end

            column = cell_index.delete('0-9')
            shared = (node_type == 's')
            next
          end
        when Nokogiri::XML::Reader::TYPE_END_ELEMENT
          if node.name == 'row'
            y << process_row(cell_map)
          end
          next
        end

        if node.value
          value = (shared ? string_lookup(value.to_i) : value)
          case cell_type
            when :datetime
            when :time
            when :date
              value = (DateTime.new(1899,12,30) + node.value.to_f)
            when :percentage # ? TODO
            when :float
              value = node.value.to_f
            else
              # leave as string
          end
          cell_type = nil
          cell_map[column] = value
        end
      end
    end
  end

  def process_row(cell_map)
    max = cell_map.keys.map {|c| self.class.column_name_to_index c }.max
    row = []
    self.class.column_names[0..max].each do |col|
      if self.class.column_name_to_index(col) > max
        break
      else
        row << cell_map[col]
      end
    end
    row
  end



  # Returns A to ZZZ.
  def self.column_names
    if @column_names
      @column_names
    else
      proc = Proc.new do |prev|
        ("#{prev}A".."#{prev}Z").to_a
      end
      x = proc.call("")
      y = x.map(&proc).flatten
      z = y.map(&proc).flatten
      @column_names = x + y + z
    end
  end

  def self.column_name_to_index(name)
    if not @column_names_to_indices
      @column_names_to_indices = {}
      self.column_names.each_with_index do |name, i|
        @column_names_to_indices[name] = i
      end
    end
    @column_names_to_indices[name]
  end

  def row_count
    if defined? @row_count
      @row_count
    elsif @file
      @file.rewind
      Nokogiri::XML::Reader(@file).each do |node|
        if node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT
          case node.name
          when 'dimension'
            if ref = node.attributes["ref"]
              break @row_count = ref.scan(/\d+$/).first.to_i
            end
          when 'sheetData'
            break @row_count = nil
          end
        end
      end
    end
  end

  private
  def path
    "xl/worksheets/sheet#{@index}.xml"
  end

end
