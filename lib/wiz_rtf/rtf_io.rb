# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  class RtfIO
    def initialize(io = nil)
      @io = io || StringIO.new
    end

    def write(txt)
      @io.write txt
    end

    def cmd(name, value = nil)
      @io.write '\\'
      @io.write name
      @io.write value if value
    end

    def txt(str)
      str = str.to_s
      str = str.gsub("{", "\\{").gsub("}", "\\}").gsub("\\", "\\\\")
      str = str.encode("UTF-16LE", :undef=>:replace).each_codepoint.map {|n| n < 128 ? n.chr : "\\u#{n}\\'3f"}.join('')
      @io.write ' '
      @io.write str
    end

    def group
      @io.write '{'
      yield if block_given?
      @io.write '}'
    end

    def delimit
      yield if block_given?
      @io.write ';'
    end
  end
end
