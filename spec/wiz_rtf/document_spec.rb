# encoding: utf-8
require 'spec_helper'

describe 'WizRtf::Document' do

  it 'is example of document ' do
    doc = WizRtf::Document.new do
      text "学生综合素质报告", :align => :center, 'font-size' => 48
      image('h:\eahey.png')
      page_break
      text "A Table Demo"
      table [[{content:'e',rowspan:4},{content:'4',rowspan:4},1,{content:'1',colspan:2}],
             [{content:'4',rowspan:3,colspan:2},8],[11]], column_widths:{1=>100,2 => 100,3 => 50,4 => 50,5 => 50} do
        add_row [1]
      end
    end
    doc.save('c:\text.rtf')
  end
end