# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  class Document
    def initialize(&block)
      @fonts = []
      @colors = []
      @parts = []
      font 0, 'fswiss', 'Arial', 0, 2
      font 1, 'fmodern', 'Courier New', 0, 1
      font 2, 'fnil', '宋体', 2, 2
      color 0, 0, 0
      color 255, 0, 0
      color 255, 0, 255
      block.arity<1 ? self.instance_eval(&block) : block.call(self) if block_given?
    end

    def head

    end

    def font(num, family, name, character_set = 0, prq = 2)
      @fonts << WizRtf::Font.new(num, family, name, character_set, prq)
    end

    def color(red, green, blue)
      @colors << WizRtf::Color.new(red, green, blue)
    end

    def text(str, styles = {:align => :left})
      @parts << WizRtf::Text.new(str, styles)
    end

    def image(file)
      @parts << WizRtf::Image.new(file)
    end

    def table(rows = [],options = {}, &block)
      @parts << WizRtf::Table.new(rows, options, &block)
    end

    def line_break
      @parts << WizRtf::Cmd.new(:par)
    end

    def page_break
      @parts << WizRtf::Cmd.new(:page)
    end

    def render(io)
      io.group do
        io.cmd :rtf, 1
        io.cmd :ansi
        io.cmd :ansicpg, 2052
        io.cmd :deff, 0
        io.group do
          io.cmd :fonttbl
          @fonts.each do |font|
            font.render(io)
          end
        end
        io.group do
          io.cmd :colortbl
          io.delimit
          @colors.each do |color|
            color.render(io)
          end
        end
        @parts.each do |part|
          part.render(io)
        end
      end
    end

    def save(file)
      File.open(file, 'w') { |file| render(WizRtf::RtfIO.new(file)) }
    end

  end
end
