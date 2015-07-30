# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  class Color
    RED = '#FF0000'
    YELLOW = '#FFFF00'
    LIME = '#00FF00'
    CYAN = '#00FFFF'
    BLUE = '#0000FF'
    MAGENTA = '#FF00FF'
    MAROON = '#800000'
    OLIVE = '#808000'
    GREEN = '#008000'
    TEAL = '#008080'
    NAVY = '#000080'
    PURPLE = '#800080'
    WHITE = '#FFFFFF'
    SILVER = '#C0C0C0'
    GRAY = '#808080'
    BLACK = '#000000'

    def initialize(*rgb)
      case rgb.size
        when 1
          from_rgb_hex(rgb.first)
        when 3
          from_rgb(*rgb)
      end
    end

    def from_rgb_hex(color)
      color = '#%.6x' % color if color.is_a? Integer
      rgb = color[1,7].scan(/.{2}/).map{ |c| c.to_i(16) }
      from_rgb(*rgb)
    end

    def from_rgb(*rgb)
      @red, @green, @blue = rgb
    end

    def to_rgb
      [@red, @green, @blue]
    end

    def to_rgb_hex
      "#" + to_rgb.map {|c| "%02X" % c }.join
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
