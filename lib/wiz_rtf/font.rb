# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  class Font

    # 新细明体：PMingLiU
    # 细明体：MingLiU
    # 标楷体：DFKai-SB
    # 黑体：SimHei
    # 宋体：SimSun
    # 新宋体：NSimSun
    # 仿宋：FangSong
    # 楷体：KaiTi
    # 微软正黑体：Microsoft JhengHei
    # 微软雅黑体：Microsoft YaHei
    FAMILIES = [
        {family:'froman', name:'Times New Roman', character:0, prq:2},
        {family:'fnil', name:'KaiTi', character:0, prq:2},
        {family:'fnil', name:'FangSong', character:0, prq:2},
        {family:'fmodern', name:'NSimSun', character:134, prq:1},
        {family:'fswiss', name:'Microsoft YaHei', character:134, prq:2}
    ]



    attr_accessor :name, :num
    def initialize(num, family, name, character_set = 0, prq = 2)
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

