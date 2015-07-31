# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  # = the Rtf Document Table.
  class Table
    DEFAULT_COLUMN_WIDTH = 40
    attr_accessor :row_spans, :column_widths

    # Creates a new Table
    # * +rows+ - a table can be thought of as consisting of rows and columns.
    # == Options:
    # * +column_widths+ - sets the widths of the Columns.
    # == Example:
    #
    #   WizRtf::Table.new([
    #       [{content: WizRtf::Image.new('h:\eahey.png'),rowspan:4},{content:'4',rowspan:4},1,{content:'1',colspan:2}],
    #       [{content:'4',rowspan:3,colspan:2},8],[11]
    #     ], column_widths:{1=>100,2 => 100,3 => 50,4 => 50,5 => 50}) do
    #     add_row [1]
    #   end
    #
    def initialize(rows = [], options = {}, &block)
      @rows = []
      @row_spans = {}
      @column_widths = options[:column_widths] || DEFAULT_COLUMN_WIDTH
      rows.each_index do |index|
        add_row rows[index]
      end
      block.arity<1 ? self.instance_eval(&block) : block.call(self) if block_given?
    end

    # Add The Cells Array of the Row.
    # * +cells+ - the cells array.
    # == Example:
    #
    #   add_row [{content:'4',rowspan:3,colspan:2},8]
    #
    def add_row(cells = [])
      @rows << WizRtf::Row.new(self, cells)
    end

    # Outputs the Partial Rtf Document to a Generic Stream as a Rich Text Format (RTF).
    # * +io+ - The Generic IO to Output the RTF Document.
    def render(io)
      @rows.each do |row|
        row.render(io)
      end
    end

  end
end