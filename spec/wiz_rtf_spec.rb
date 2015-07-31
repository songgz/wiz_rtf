require 'spec_helper'

describe WizRtf do
  it 'has a version number' do
    expect(WizRtf::VERSION).not_to be nil
  end

  it 'is example of document ' do
    doc = WizRtf::Document.new default_font:'NSimSun' do
      text "学生综合素质报告", 'text-align' => :center, 'font-family' => 'Microsoft YaHei', 'font-size' => 48, 'font-bold' => true, 'font-italic' => true, 'font-underline' => true
      image('h:\eahey.png')
      page_break
      text "A Table Demo", 'foreground-color' => WizRtf::Color::RED, 'background-color' => '#0f00ff'
      table [[{content: WizRtf::Image.new('h:\eahey.png'),rowspan:4},{content:'4',rowspan:4},1,{content:'1',colspan:2}],
             [{content:'4',rowspan:3,colspan:2},8],[11]], column_widths:{1=>100,2 => 100,3 => 50,4 => 50,5 => 50} do
        add_row [1]
      end
    end
    doc.save('c:\text.rtf')
  end

  it 'is example of document for table' do
    doc = WizRtf::Document.new default_font:'NSimSun' do
      text "学生成绩单", 'foreground-color' => WizRtf::Color::RED, 'text-align' => :center, 'font-size' => 24
      table [
                [{content:'姓名', rowspan:3},{content:'语文', colspan:5}],
                ['期中','期末',{content:'平时', colspan:3}],
                ['期中','期末','课堂','实践','作业']
            ], column_widths:{1=>100,2 => 100,3 => 50,4 => 50,5 => 50}
    end
    doc.save('c:\table.rtf')
  end

end
