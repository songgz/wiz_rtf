# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  # = Rtf Document
  #
  # Creates a new Rtf document specifing the format of the pages.
  # == Example:
  #
  #   doc = WizRtf::Document.new default_font:'NSimSun' do
  #     text "A Example of Rtf Document", 'text-align' => :center, 'font-family' => 'Microsoft YaHei', 'font-size' => 48, 'font-bold' => true, 'font-italic' => true, 'font-underline' => true
  #     image('h:\eahey.png')
  #     page_break
  #     text "A Table Demo", 'foreground-color' => WizRtf::Color::RED, 'background-color' => '#0f00ff'
  #     table [[{content: WizRtf::Image.new('h:\eahey.png'),rowspan:4},{content:'4',rowspan:4},1,{content:'1',colspan:2}],
  #            [{content:'4',rowspan:3,colspan:2},8],[11]], column_widths:{1=>100,2 => 100,3 => 50,4 => 50,5 => 50} do
  #       add_row [1]
  #     end
  #   end
  #   doc.save('c:\text.rtf')
  #
  class Document
    def initialize(options = {}, &block)
      @fonts = []
      @colors = []
      @parts = []
      font options[:default_font] || 'Arial'
      block.arity<1 ? self.instance_eval(&block) : block.call(self) if block_given?
    end

    # Outputs the Complete Rtf Document to a Generic Stream as a Rich Text Format (RTF)
    # * +io+ - The Generic IO to Output the RTF Document
    def render(io)
      io.group do
        io.cmd :rtf, 1
        io.cmd :ansi
        io.cmd :ansicpg, 2052
        io.cmd :deff, 0
        io.group do
          io.cmd :fonttbl
          @fonts.each_with_index do |font, index|
            io.group do
              io.delimit do
                io.cmd :f, index
                io.cmd font.family
                io.cmd :fprq, font.pitch
                io.cmd :fcharset, font.character_set
                io.write ' '
                io.write font.name
              end
            end
          end
        end
        io.group do
          io.cmd :colortbl
          io.delimit
          @colors.each do |color|
            io.delimit do
              io.cmd :red, color.red
              io.cmd :green, color.green
              io.cmd :blue, color.blue
            end
          end
        end
        @parts.each do |part|
          part.render(io)
        end
      end
    end

    # Outputs the complete Rtf Document to a file as a Rich Text Format (RTF)
    # * +file+ - file path and filename.
    def save(file)
      File.open(file, 'w') { |file| render(WizRtf::RtfIO.new(file)) }
    end

    def head(&block)
      block.arity<1 ? self.instance_eval(&block) : block.call(self) if block_given?
    end

    # Sets the Font for the text.
    def font(name, family = nil,  character_set = 0, pitch = 2)
      unless index = @fonts.index {|f| f.name == name}
        index = @fonts.size
        opts = WizRtf::Font::FONTS.detect {|f| f[:name] == name}
        if opts
          @fonts << WizRtf::Font.new(opts[:name], opts[:family], opts[:character], opts[:pitch])
        else
          @fonts << WizRtf::Font.new(name, family, character_set, pitch)
        end
      end
      index
    end

    # Sets the color for the text.
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

    # This will add a string of +str+ to the document, starting at the
    # current drawing position.
    # == Styles:
    # * +text-align+ - sets the horizontal alignment of the text. optional values: +:left+, +:center+, +:right+
    # * +font-family+ - set the font family of the text.
    #    - optional values:
    #         'Arial', 'Arial Black', 'Arial Narrow','Bitstream Vera Sans Mono',
    #         'Bitstream Vera Sans','Bitstream Vera Serif','Book Antiqua','Bookman Old Style','Castellar','Century Gothic',
    #         'Comic Sans MS','Courier New','Franklin Gothic Medium','Garamond','Georgia','Haettenschweiler','Impact','Lucida Console'
    #         'Lucida Sans Unicode','Microsoft Sans Serif','Monotype Corsiva','Palatino Linotype','Papyrus','Sylfaen','Symbol'
    #         'Tahoma','Times New Roman','Trebuchet MS','Verdana'.
    # * +font-size+ - set font size of the text.
    # * +font-bold+ - setting the value true for bold of the text.
    # * +font-italic+ - setting the value true for italic of the text.
    # * +font-underline+ - setting the value true for underline of the text.
    # == Example:
    #
    #  text "A Example of Rtf Document", 'text-align' => :center, 'font-family' => 'Microsoft YaHei', 'font-size' => 48, 'font-bold' => true, 'font-italic' => true, 'font-underline' => true
    #
    def text(str, styles = {})
      styles['foreground-color'] = color(styles['foreground-color']) if styles['foreground-color']
      styles['background-color'] = color(styles['background-color']) if styles['background-color']
      styles['font-family'] = font(styles['font-family']) if styles['font-family']
      @parts << WizRtf::Text.new(str, styles)
    end

    # Puts a image into the current position within the document.
    # * +file+ - image file path and filename.
    # == Example:
    #
    #   image('h:\eahey.png')
    #
    def image(file)
      @parts << WizRtf::Image.new(file)
    end

    # Creates a new Table
    # * +rows+ -  a table can be thought of as consisting of rows and columns.
    # == Options:
    # * +column_widths+ - sets the widths of the Columns.
    # == Example:
    #
    #   table [
    #       [{content: WizRtf::Image.new('h:\eahey.png'),rowspan:4},{content:'4',rowspan:4},1,{content:'1',colspan:2}],
    #       [{content:'4',rowspan:3,colspan:2},8],[11]
    #     ], column_widths:{1=>100,2 => 100,3 => 50,4 => 50,5 => 50} do
    #     add_row [1]
    #   end
    #
    def table(rows = [],options = {}, &block)
      @parts << WizRtf::Table.new(rows, options, &block)
    end

    # Writes a new line.
    def line_break
      @parts << WizRtf::Cmd.new(:par)
    end

    # Writes a page interruption (new page)
    def page_break
      @parts << WizRtf::Cmd.new(:page)
    end

  end
end
