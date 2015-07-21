# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  class Table
    DEFAULT_COLUMN_WIDTH = 40
    attr_accessor :row_spans, :column_widths

    def initialize(rows = [], options = {}, &block)
      @rows = []
      @row_spans = {}
      @column_widths = options[:column_widths] || DEFAULT_COLUMN_WIDTH
      rows.each_index do |index|
        add_row rows[index]
      end
      block.arity<1 ? self.instance_eval(&block) : block.call(self) if block_given?
    end

    def add_row(cells = [])
      @rows << WizRtf::Row.new(self, cells)
    end

    def render(io)
      @rows.each do |row|
        row.render(io)
      end
    end

  end
end