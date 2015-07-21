# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  class Cmd
    def initialize(name, value = nil)
      @name = name
      @value = value
    end

    def render(io)
      io.cmd @name, @value
    end
  end
end
