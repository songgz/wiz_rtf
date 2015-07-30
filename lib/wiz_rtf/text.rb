# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf

  # = Represents Rtf text.
  class Text
    TEXT_ALIGN_MAP = {left:'ql',center:'qc',right:'qr'}

    # creates a text on the document.
    # <tt>:options</tt>:: set the text alignment.
    # * <tt>:text-align</tt>:: sets the horizontal alignment of the text.
    #  - optional values:: <tt>:left</tt>, <tt>:center</tt>, <tt>:right</tt>
    # * <tt>:font-family</tt>:: set the font family of the text.
    #  - optional values::
    # * <tt>:font-size</tt>:: set font size of the text.
    # * <tt>:font-bold</tt>:: setting the value true for bold of the text.
    # * <tt>:font-italic</tt>:: setting the value true for italic of the text.
    # * <tt>:font-underline</tt>:: setting the value true for underline of the text.
    # Example:
    # WizRtf::Text.new("A Example of Rtf Document", 'text-align' => :center, 'font-family' => 'Microsoft YaHei', 'font-size' => 48, 'font-bold' => true, 'font-italic' => true, 'font-underline' => true)
    def initialize(str = '', styles = {})
      @str = str
      @styles = {'text-align' => :left, 'font-family' => 0, 'font-size' => 24, 'font-bold' => false, 'font-italic' => false, 'font-underline' => false, 'foreground-color' => 0, 'background-color' => 0 }.merge(styles)
    end

    # Outputs the Partial Rtf Document to a Generic Stream as a Rich Text Format (RTF).
    # <tt>:io</tt>:: The Generic IO to Output the RTF Document.
    def render(io)
      io.group do
        io.cmd :pard
        io.cmd TEXT_ALIGN_MAP[@styles['text-align']]
        io.cmd :f, @styles['font-family']
        io.cmd :fs, @styles['font-size']
        io.cmd @styles['font-bold'] ? 'b' : 'b0'
        io.cmd @styles['font-italic'] ? 'i' : 'i0'
        io.cmd @styles['font-underline'] ? 'ul' : 'ulnone'
        io.cmd :cf, @styles['foreground-color']
        io.cmd :cb, @styles['background-color']
        io.cmd :chcfpat, @styles['foreground-color']
        io.cmd :chcbpat, @styles['background-color']
        io.txt @str
        io.cmd :par
      end
    end
  end
end

