# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  class Font
    def initialize(num, family, name, character_set = 0, prq = 2)
      @num = num
      @family = family
      @name = name
      @character_set = character_set
      @prq = prq
    end

    def render(io)
      io.group do
        io.delimit do
          io.cmd :f, @num
          io.cmd @family
          io.cmd :fprq, @prq
          io.cmd :fcharset, @character_set
          io.write ' '
          io.write @name
        end
      end
    end
  end
end

