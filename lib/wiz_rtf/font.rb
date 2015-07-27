# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  class Font
    FAMILIES = {
        default: 'fnil',
        roman: 'froman',
        swiss: 'fswiss',
        fixed_pitch: 'fmodern',
        script: 'fscript',
        decorative: 'fdecor',
        technical: 'ftech',
        bidirectional: 'fbidi'
    }

    CHARACTER_SET = {
        ansi: 0,
        default: 1,
        symbol: 2,
        invalid: 3,
        mac: 77,
        shiftJis: 128,
        hangul: 129,
        johab: 130,
        gb2312: 134,
        big5: 136,
        greek: 161,
        turkish: 162,
        vietnamese: 163,
        hebrew: 177,
        arabic: 178,
        arabicTraditional: 179,
        arabic_user: 180,
        hebrew_user: 181,
        baltic: 186,
        russian: 204,
        thai: 222,
        eastern_european: 238,
        pc437: 254,
        oem: 255
    }

    FONTS = [
        {family:'fswiss', name:'Arial', character:0, prq:2},
        {family:'froman', name:'Courier New', character:0, prq:1},
        {family:'froman', name:'Times New Roman', character:0, prq:2},
        {family:'fnil', name:'SimSun', character:134, prq:2},
        {family:'fmodern', name:'KaiTi', character:134, prq:1},
        {family:'fnil', name:'FangSong', character:134, prq:1},
        {family:'fmodern', name:'SimHei', character:134, prq:1},
        {family:'fmodern', name:'NSimSun', character:134, prq:1},
        {family:'fswiss', name:'Microsoft YaHei', character:134, prq:2}
    ]

    attr_accessor :name, :num
    def initialize(num, name, family = 'fnil', character_set = 0, prq = 2)
      @num = num
      @family = family if family
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

