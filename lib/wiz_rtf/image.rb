# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf
  class Image
    JPEG_SOF_BLOCKS = [0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF]
    PIC_TYPE = {png: :pngblip, jpg: :jpegblip, bmp: :pngblip, gif: :pngblip}

    def initialize(file)
      begin
        @img = IO.binread(file)
        @width, @height = self.size
      rescue  Exception => e
        STDERR.puts "** error parsing #{file}: #{e.inspect}"
      end
    end

    def type
      png = Regexp.new("\x89PNG".force_encoding("binary"))
      jpg = Regexp.new("\xff\xd8\xff\xe0\x00\x10JFIF".force_encoding("binary"))
      jpg2 = Regexp.new("\xff\xd8\xff\xe1(.*){2}Exif".force_encoding("binary"))

      @type = case @img
                when /^GIF8/
                  :gif
                when /^#{png}/
                  :png
                when /^#{jpg}/
                  :jpg
                when /^#{jpg2}/
                  :jpg
                when /^BM/
                  :bmp
                else
                  :unknown
              end
    end

    def size
      case self.type
        when :gif
          @img[6..10].unpack('SS')
        when :png
          @img[16..24].unpack('NN')
        when :bmp
          d = @img[14..28]
          d.unpack('C')[0] == 40 ? d[4..-1].unpack('LL') : d[4..8].unpack('SS')
        when :jpg
          d = StringIO.new(@img)
          section_marker = 0xff # Section marker.
          d.seek(2) # Skip the first two bytes of JPEG identifier.
          loop do
            marker, code, length = d.read(4).unpack('CCn')
            fail "JPEG marker not found!" if marker != section_marker
            if JPEG_SOF_BLOCKS.include?(code)
              #@bits, @height, @width, @channels = d.read(6).unpack("CnnC")
              return d.read(6).unpack('CnnC')[1..2].reverse
            end
            d.seek(length - 2, IO::SEEK_CUR)
          end
      end
    end

    def render(io)
      io.group do
        io.cmd '*'
        io.cmd 'shppict'
        io.group do
          io.cmd :pict
          io.cmd PIC_TYPE[@type]
          io.cmd :picscalex, 99
          io.cmd :picscaley, 99
          io.cmd :picw, @width
          io.cmd :pich, @height
          io.write "\n"
          @img.each_byte do |b|
            hex_str = b.to_s(16)
            hex_str.insert(0,'0') if hex_str.length == 1
            io.write hex_str
          end
          io.write "\n"
        end
      end if @img
    end
  end
end