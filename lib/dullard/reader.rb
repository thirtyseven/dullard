require 'zip/zipfilesystem'
require 'nokogiri'
require 'pp'

module Dullard; end

class Dullard::Workbook
  # Code borrowed from Roo (https://github.com/hmcgowan/roo/blob/master/lib/roo/excelx.rb)
  # Some additional formats added by Paul Hendryx that are common in LibreOffice.
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
    @zipfs = Zip::ZipFile.open(@file)
    @user_defined_formats = user_defined_formats
    read_styles
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

  def read_styles
    doc = Nokogiri::XML(@zipfs.file.open("xl/styles.xml"))
    
    @numFmts = {}
    @cellXfs = []
    fonts = []
    
    doc.css('/styleSheet/numFmts/numFmt').each do |numFmt|
      numFmtId = numFmt.attributes['numFmtId'].value.to_i
      formatCode = numFmt.attributes['formatCode'].value
      @numFmts[numFmtId] = formatCode
    end

    doc.css('/styleSheet/cellXfs/xf').each do |xf|
      numFmtId = xf.attributes['numFmtId'].value.to_i
      @cellXfs << numFmtId
    end

    return @numFmts, @cellXfs
  end

  
  # Code borrowed from Roo (https://github.com/hmcgowan/roo/blob/master/lib/roo/excelx.rb)
  # convert internal excelx attribute to a format
  def attribute2format(s)
    id = @cellXfs[s.to_i].to_i
    result = @numFmts[id]

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
    @file = @workbook.zipfs.file.open(path) if @workbook.zipfs.file.exist?(path)
  end

  def string_lookup(i)
    @workbook.string_table[i]
  end

  def row_count
    return rows_size
  end

  def rows
    Enumerator.new(rows_size) do |y|
      next unless @file
      @file.rewind
      shared = false
      row = nil
      column = nil
      cell_type = nil
      Nokogiri::XML::Reader(@file).each do |node|
        case node.node_type
        when Nokogiri::XML::Reader::TYPE_ELEMENT
          case node.name
          when "row"
            row = []
            column = 0
            next
          when "c"
            if node.attributes['t'] != 's' && node.attributes['t'] != 'b'
              cell_format_index = node.attributes['s'].to_i
              cell_type = @workbook.format2type(@workbook.attribute2format(cell_format_index))
            end

            rcolumn = node.attributes["r"]
            if rcolumn
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
        value = node.value

        if value
          case cell_type
            when :datetime
            when :time
            when :date
              value = (Date.parse('1900-01-01') + value.to_f)
            when :percentage # ? TODO
            when :float
              value = value.to_f
            else
              # leave as string
          end
          cell_type = nil

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
      proc = Proc.new do |prev|
        ("#{prev}A".."#{prev}Z").to_a
      end
      x = proc.call("")
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
