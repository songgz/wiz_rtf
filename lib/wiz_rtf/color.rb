# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  class Color

    def initialize(red, green, blue)
      @red = red
      @green = green
      @blue = blue
    end

    def render(io)
      io.delimit do
        io.cmd :red, @red
        io.cmd :green, @green
        io.cmd :blue, @blue
      end
    end
  end
end
