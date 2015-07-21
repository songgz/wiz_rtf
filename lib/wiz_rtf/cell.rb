# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  class Cell
    attr_accessor :colspan, :rowspan, :content, :v_merge, :right_width

    def initialize(cell)
      unless cell.is_a?(Hash)
        @colspan = 1
        @rowspan = 1
        @content = cell
      else
        @colspan = cell[:colspan] || 1
        @rowspan = cell[:rowspan] || 1
        @content = cell[:content] || ''
      end
    end

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
      io.txt content
      io.cmd :cell
    end

  end
end

