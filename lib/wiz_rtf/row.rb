# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf

  # =  Represents a table row.
  class Row

    # This is the constructor for the Row class.
    # * +table+ - A reference to table that owns the row.
    # * +cells+ - The number of cells that the row will contain.
    def initialize(table, cells = [])
      @table = table
      @cells = []
      @col_offset = 1
      @right_width = 0
      cells.each do |cell|
        add_cell cell
      end
    end

    # add a Cell object to the Cells array.
    # * +cell+ - a Cell object.
    # * +merge+ - is merges the specified table cells.
    def add_cell(cell, merge = false)
      add_cell('', true) if !merge && row_spanned?(@col_offset)

      c = WizRtf::Cell.new(cell)
      @table.row_spans[@col_offset] = c if c.rowspan > 1

      if c.rowspan > 1
        c.v_merge = :clvmgf
      elsif row_spanned? @col_offset
        c.v_merge = :clvmrg
        @table.row_spans[@col_offset].rowspan -= 1
        c.colspan = @table.row_spans[@col_offset].colspan || c.colspan
      end
      c.colspan.times do
        @right_width += column_width(@col_offset)
        @col_offset += 1
      end
      c.right_width = @right_width

      @cells << c

      add_cell('', true) if row_spanned?(@col_offset)
    end

    def row_spanned?(offset)
      @table.row_spans[offset] && @table.row_spans[offset].rowspan > 1
    end

    def column_width(offset)
      return 20 * (@table.column_widths[offset] || WizRtf::Table::DEFAULT_COLUMN_WIDTH) if @table.column_widths.is_a?(Hash)
      @table.column_widths * 20
    end

    # Outputs the Partial Rtf Document to a Generic Stream as a Rich Text Format (RTF).
    # * +io+ - The Generic IO to Output the RTF Document.
    def render(io)
      io.cmd :trowd
      io.cmd :trbrdrt
      io.cmd :brdrs
      io.cmd :brdrw10
      io.cmd :trbrdrl
      io.cmd :brdrs
      io.cmd :brdrw10
      io.cmd :trbrdrb
      io.cmd :brdrs
      io.cmd :brdrw10
      io.cmd :trbrdrr
      io.cmd :brdrs
      io.cmd :brdrw10
      io.cmd :trautofit1
      io.cmd :intbl
      @cells.each do |c|
        c.render(io)
      end
      io.cmd :row
    end
  end
end

