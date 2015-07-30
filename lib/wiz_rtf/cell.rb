# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf

  # = Represents a table cell.
  class Cell
    attr_accessor :colspan, :rowspan, :content, :v_merge, :right_width

    # This is the constructor for the Cell class.
    # * +cell+ - optional values:: number, string, symbol, hash.
    # == Example:
    # # WizRtf::Cell.new({content:'4', rowspan:3, colspan:2})
    def initialize(cell)
      if cell.is_a?(Hash)
        @colspan = cell[:colspan] || 1
        @rowspan = cell[:rowspan] || 1
        @content = cell[:content] || ''
      else
        @colspan = 1
        @rowspan = 1
        @content = cell
      end
    end

    # Outputs the Partial Rtf Document to a Generic Stream as a Rich Text Format (RTF).
    # * +io+ - The Generic IO to Output the RTF Document.
    def render(io)
      io.cmd :celld
      io.cmd :clbrdrt
      io.cmd :brdrs
      io.cmd :brdrw10
      io.cmd :clbrdrl
      io.cmd :brdrs
      io.cmd :brdrw10
      io.cmd :clbrdrb
      io.cmd :brdrs
      io.cmd :brdrw10
      io.cmd :clbrdrr
      io.cmd :brdrs
      io.cmd :brdrw10
      io.cmd  v_merge if v_merge
      io.cmd :cellx, right_width
      contents = [@content] unless @content.is_a?(Array)
      contents.each do |c|
        if c.respond_to?(:render)
          c.render(io)
        else
          io.txt c
        end
      end
      io.cmd :cell
    end

  end
end

