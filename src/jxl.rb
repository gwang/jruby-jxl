#---------------------------------------------------------------------------------------
# jxl.rb
#
# This provides interface to the Jxl lib for processing xls files.
#
# Tested: under Windows 7 using JRuby 1.7.3, Java 7, and jxl 2.6.10.lib
#
# Copyright: Gang Wang @ 2013
#----------------------------------------------------------------------------------------
require 'java'

class String
  def get_attr_symbol
    self.upcase.to_sym
  end
end

module Jxl
  require '../lib/jxl-2.6.10.jar'
  include_package 'jxl'
  include_package 'jxl.write'
  include_package 'jxl.format'

  JFile = java.io.File

  class Style
    attr_accessor :config
    attr_accessor :fmt, :allowed, :font

    # allowed inputs looks like the following:
    # ( :bkcolor => 'yellow',
    #   :alignment=>'center',
    #   :wrap => true,
    #   :font => { :name => 'arial',
    #              :size => 12,
    #              :underline_style=>'single',
    #              :bold => true
    #            }
    #   :border => {:where => 'all',
    #               :line_style => "medium",
    #               :wrap => true}
    # )
    # for all the allowed values, please run the #info method
    def initialize(params)
      @allowed = {}
      @allowed[:color]             = Jxl::Colour.constants
      @allowed[:border]            = Jxl::Border.constants
      @allowed[:border_line_style] = Jxl::BorderLineStyle.constants
      @allowed[:underline_style]   = Jxl::UnderlineStyle.constants
      @allowed[:font]              = Jxl::WritableFont.constants

      # somehow the following line does not work
      #@allowed[:alignment]         = Java::JxlFormat::Alignment.constants
      @allowed[:alignment] = [:GENERAL, :LEFT, :CENTRE, :RIGHT, :FILL, :JUSTIFY]

      # info
      @config = params
      configure_format
    end

    def info
      @allowed.each { |k,v|
        puts "allowed #{k}(s) = #{v}"
      }
    end

    def configuration
      @config
    end

    def configure_format
      @font = configure_font
      @fmt = @font.nil? ? Jxl::WritableCellFormat.new() : Jxl::WritableCellFormat.new(@font)
      set_format_parameter(Jxl::Colour,    'setBackground', [:bkcolor])

      #set_format_parameter(Java::JxlFormat::Alignment, 'setAlignment',  [:alignment])
      # hacking !!!
      x = @allowed[:alignment].index(@config[:alignment].get_attr_symbol)
      if ( x != nil)
        a = Java::JxlFormat::Alignment.getAlignment(x)
        @fmt.setAlignment(a)
      else
        puts "Unsupported alignment specification: #{@config[:alignment]}. Nothing happens."
      end
      set_format_parameter(Jxl::Alignment, 'setWrap',       [:wrap])

      #handling border settings
      if @config[:border]
        @config[:border] = {:where => 'all',
                            :line_style=>'single',
                            :color => 'black'}.merge(@config[:border])
        @fmt.setBorder( get_format_constant(Jxl::Border,          [:border, :where]),
                        get_format_constant(Jxl::BorderLineStyle, [:border, :line_style]),
                        get_format_constant(Jxl::Colour,          [:border, :color] ))
      end
    end

    def configure_font
      ret = nil
      unless @config[:font] == nil
        @config[:font] = {:name=>'arial',
                          :size=>11,
                          :underline_style=>'no_underline',
                          :italic=>false,
                          :color=>'black',
                          :bold=>false}.merge(@config[:font])

        x = get_format_constant(Jxl::WritableFont, [:font, :name])
        ret = Jxl::WritableFont.new(x, @config[:font][:size])

        x = get_format_constant(Jxl::UnderlineStyle, [:font, :underline_style])
        ret.setUnderlineStyle(x)

        if @config[:font][:bold] == true
          ret.setBoldStyle(Jxl::WritableFont.const_get :BOLD)
        else
          ret.setBoldStyle(Jxl::WritableFont.const_get :NO_BOLD)
        end

        ret.setItalic(@config[:font][:italic])
        ret.setColour(get_format_constant(Jxl::Colour, [:font, :color]))
      end
      ret
    end

    def get_format_constant(attribute, hash_names)
      return nil if hash_names.size == 0
      v = @config
      hash_names.each do |n|
        if v[n] != nil
          v = v[n]
        else
          puts "broken with : attribute = #{attribute}, hash_names = #{hash_names}"
          v = nil
          break
        end
      end
      # puts "in get_format_constant: #{v}"
      return v.nil? ? nil : ((v.instance_of?(String) || v.instance_of?(Symbol)) ? attribute.const_get(v.get_attr_symbol) : v)
    end

    def set_format_parameter(attribute, function_name, hash_names)
        x = get_format_constant(attribute, hash_names)
        @fmt.send(function_name, x) if x != nil
    end
  end

  module WorkbookMixin
    def sheet(sheet_name, position = 0, &block)
      sheet = getSheet(sheet_name)
      unless sheet
        sheet = createSheet(sheet_name, position)
      end
      sheet.extend SheetMixin
      yield sheet if block_given?
      sheet
    end

    def all_sheets
      sheet_names = getSheets()
      if block_given?
        sheet_names.each do |s|
          yield getSheet(s)
        end
      else
        return sheet_names.map {|s| s.getName()}
      end
    end
  end

  module SheetMixin
    attr_accessor :merged_cells_array

    def merged_cells_to_array()
      ret = []
      puts "getting merged cells"
      cells = getMergedCells()
      puts "merged cells = #{cells}, size = #{cells.size}"
      cells.each do |c|
         ret << [c.topLeft.row,     c.topLeft.column,
                 c.bottomRight.row, c.bottomRight.column]
      end
      self.merged_cells_array = ret
      ret
    end

    def is_part_of_merged_cell?(cell)
      self.merged_cells_array.each do |mc|
        top, left, bottom, right= mc
        row, column = cell.row, cell.column
        if ( row >= top and row <= bottom and column >= left and column <= right)
          return (row == top and column == left) ? false : true
        end
      end
      return false
    end

    {
      :label => Jxl::Label,
      :number => Jxl::Number,
    }.each do |name, klass|
      define_method name do |*args|
        for i in (0..args.size)
          args[i] = args[i].fmt if args[i].class == Style
        end
        cell = klass.new(*args)
        yield cell if block_given?
        addCell cell
        #automatic resize the cell to fit content
        x = cell.getColumn
        # cv = getColumnView(x)
        # cv.setAutosize(true)
        # setColumnView(x, cv)
        setColumnView(x, 30)
      end
    end
  end # end of module SheetMixin

  module CellMixin
  end

  class <<self
    def create(file_name)
      file = JFile.new(file_name)
      workbook = Jxl::Workbook.createWorkbook(file)
      workbook.extend WorkbookMixin
      yield workbook if block_given?
      workbook.write
    ensure
      workbook.close if defined?(workbook) && workbook
    end

    def open(file_name)
      file = JFile.new(file_name)
      workbook = Jxl::Workbook.getWorkbook(file)
      workbook.extend WorkbookMixin
      yield workbook if block_given?
    ensure
      workbook.close if defined?(workbook) && workbook
    end

  end # end of static module methods

end

#----------------- self testing code --------------------#
if __FILE__ == $0
  puts 'Test writing to an .xls file'
  require 'date'
  date = Date.today.to_s
  begin
    Jxl.create "report_#{date}.xls" do |workbook|
      header_format = Jxl::Style.new(:font => {:bold=>true,
                                               :size=>12,
                                               :underline_style => 'double'
                                               },
                                     :border => {:where => 'all',
                                                 :line_style=>'medium',
                                                 :color=>'red'
                                                 },
                                     :wrap => true,
                                     :alignment => 'right',
                                     :bkcolor => 'yellow')
      workbook.sheet("Sales #{date}") do |sheet|
        sheet.label 0, 0, "City\nName", header_format
        sheet.label 1, 0, "Total amount", header_format
        sheet.label 0, 1, "NY"
        sheet.number 1, 1, 123
      end
    end
  rescue Exception => msg
    puts "Error: #{msg}"
  end

  puts 'Test reading from an .xls file'
  #Jxl.open "report_#{date}.xls" do |workbook|
  Jxl.open "../data/test.xls" do |workbook|
    #puts workbook.sheets
    puts "sheets = #{workbook.all_sheets}"
    workbook.all_sheets.each do |n|
      puts "processing '#{n}' ..."
      s = workbook.sheet(n)
      s.extend Jxl::SheetMixin
      #puts "merged cells = #{s.getMergedCells().size}"
      s.merged_cells_to_array
      puts "ma = #{s.merged_cells_array}, size = #{s.merged_cells_array.size}"
      rownum=s.getRows
      puts "There are #{rownum} rows in this worksheet"
      0.upto(rownum-1) do |i|
        cells=s.getRow(i)
        puts "processing row #{i} with #{cells.size} cell(s)...."
        cell_values = []
        cells.each { |c|
          if !s.is_part_of_merged_cell?(c)
            cell_values << c.contents.strip
          end
        }
        puts "row content = #{cell_values}"
      end
    end
  end
end