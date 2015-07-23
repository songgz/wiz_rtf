require 'spec_helper'

describe WizRtf do
  it 'has a version number' do
    expect(WizRtf::VERSION).not_to be nil
  end

  it 'is example of document ' do
    doc = WizRtf::Document.new do
      text "学生综合素质报告", 'text-align' => :center, 'font-size' => 48, 'font-bold' => true, 'font-italic' => true, 'font-underline' => true
      image('h:\eahey.png')
      page_break
      text "A Table Demo", 'foreground-color' => 0, 'background-color' => 2
      table [[{content: WizRtf::Image.new('h:\eahey.png'),rowspan:4},{content:'4',rowspan:4},1,{content:'1',colspan:2}],
             [{content:'4',rowspan:3,colspan:2},8],[11]], column_widths:{1=>100,2 => 100,3 => 50,4 => 50,5 => 50} do
        add_row [1]
      end
    end
    doc.save('c:\text.rtf')
  end

end
