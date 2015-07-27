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
      font 'Courier New'
      block.arity<1 ? self.instance_eval(&block) : block.call(self) if block_given?
    end

    def head(&block)
      block.arity<1 ? self.instance_eval(&block) : block.call(self) if block_given?
    end

    def font(name, family = nil,  character_set = 0, prq = 2)
      unless index = @fonts.index {|f| f.name == name}
        index = @fonts.size
        opts = WizRtf::Font::FONTS.detect {|f| f[:name] == name}
        @fonts << WizRtf::Font.new(index, opts[:name], opts[:family], opts[:character], opts[:prq])
      end
      index
    end

    def color(*rgb)
      color = WizRtf::Color.new(*rgb)
      if index = @colors.index {|c| c.to_rgb_hex == color.to_rgb_hex}
        index += 1
      else
        @colors << color
        index = @colors.size
      end
      index
    end

    def text(str, styles = {})
      styles['foreground-color'] = color(styles['foreground-color']) if styles['foreground-color']
      styles['background-color'] = color(styles['background-color']) if styles['background-color']
      styles['font-family'] = font(styles['font-family']) if styles['font-family']
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
