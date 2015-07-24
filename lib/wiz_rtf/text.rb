# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  class Text
    TEXT_ALIGN_MAP = {left:'ql',center:'qc',right:'qr'}

    def initialize(str = '', styles = {})
      @str = str
      @styles = {'text-align' => :left, 'font-family' => 0, 'font-size' => 24, 'font-bold' => false, 'font-italic' => false, 'font-underline' => false, 'foreground-color' => 0, 'background-color' => 0 }.merge(styles)
    end

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

