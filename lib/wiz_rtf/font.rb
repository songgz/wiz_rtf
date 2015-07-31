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
        {name:'Arial', family:'fswiss', character:0, pitch:2, panose:'020b0604020202020204'},
        {name:'Arial Black', family:'fswiss', character:0, pitch:2, panose:'020b0a04020102020204'},
        {name:'Arial Narrow', family:'fswiss', character:0, pitch:2, panose:'020b0506020202030204'},
        {name:'Bitstream Vera Sans Mono', family:'fmodern', character:0, pitch:1, panose:'020b0609030804020204'},
        {name:'Bitstream Vera Sans', family:'fswiss', character:0, pitch:2, panose:'020b0603030804020204'},
        {name:'Bitstream Vera Serif', family:'froman', character:0, pitch:2, panose:'02060603050605020204'},
        {name:'Book Antiqua', family:'froman', character:0, pitch:2, panose:'02040602050305030304'},
        {name:'Bookman Old Style', family:'froman', character:0, pitch:2, panose:'02050604050505020204'},
        {name:'Castellar', family:'froman', character:0, pitch:2, panose:'020a0402060406010301'},
        {name:'Century Gothic', family:'fswiss', character:0, pitch:2, panose:'020b0502020202020204'},
        {name:'Comic Sans MS', family:'fscript', charater:0, pitch:2, panose:'030f0702030302020204'},
        {name:'Courier New', family:'froman', character:0, pitch:1, panose:'02070309020205020404'},
        {name:'Franklin Gothic Medium', family:'fswiss', character:0, picth:2, panose:'020b0603020102020204'},
        {name:'Garamond', family:'froman' , character:0, pitch:2, panose:'02020404030301010803'},
        {name:'Georgia', family:'froman', character:0, pitch:2, panose:'02040502050405020303'},
        {name:'Haettenschweiler', family:'fswiss', character:0, pitch:2, panose:'020b0706040902060204'},
        {name:'Impact', family:'fswiss', character:0, pitch:2, panose:'020b0806030902050204'},
        {name:'Lucida Console', family:'fmodern', character:0, pitch:1, panose:'020b0609040504020204'},
        {name:'Lucida Sans Unicode', family:'fswiss' , character:0, pitch:2, panose:'020b0602030504020204'},
        {name:'Microsoft Sans Serif', family:'fswiss', character:0, pitch:2, panose:'020b0604020202020204'},
        {name:'Monotype Corsiva', family:'fscript', character:0, pitch:2, panose:'03010101010201010101'},
        {name:'Palatino Linotype', family:'froman', character:0, pitch:2, panose:'02040502050505030304'},
        {name:'Papyrus', family:'fscript', character:0, pitch:2, panose:'03070502060502030205'},
        {name:'Sylfaen', family:'froman', character:0, pitch:2, panose:'010a0502050306030303'},
        {name:'Symbol', family:'froman', character:2, pitch:2, panose:'05050102010706020507'},
        {name:'Tahoma', family:'fswiss', character:0, pitch:2, panose:'020b0604030504040204'},
        {name:'Times New Roman', family:'froman', character:0, pitch:2, panose:'02020603050405020304'},
        {name:'Trebuchet MS', family:'fswiss', character:0, pitch:2, panose:'020b0603020202020204'},
        {name:'Verdana', family:'fswiss', character:0, pitch:2, panose:'020b0604030504040204'},
        {name:'SimSun', family:'fnil', character:134, pitch:2},
        {name:'KaiTi', family:'fmodern', character:134, pitch:1},
        {name:'FangSong', family:'fnil', character:134, pitch:1},
        {name:'SimHei', family:'fmodern', character:134, pitch:1},
        {name:'NSimSun', family:'fmodern', character:134, pitch:1},
        {name:'Microsoft YaHei', family:'fswiss', character:134, pitch:2}
    ]

    attr_accessor :name, :family, :character_set, :pitch, :panose, :alternate

    def initialize(name, family = 'fnil', character_set = 0, pitch = nil, panose = nil, alternate = nil)
      @name = name
      @family = family
      @character_set = character_set
      @pitch = pitch
      @panose = panose
      @alternate = alternate
    end

  end
end

