# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  class Text
    ALIGN_MAP = {left:'ql',center:'qc',right:'qr'}
    FONT_MAP = {'font-size' => :fs}

    def initialize(str = '', styles = {})
      @str = str
      @styles = {:align => :left,'font-size' => 24}.merge(styles)
    end

    def render(io)
      io.group do
        io.cmd :pard
        io.cmd ALIGN_MAP[@styles[:align]]
        io.cmd FONT_MAP[@styles['font-size']]
        io.txt @str
        io.cmd :par
      end
    end
  end
end

