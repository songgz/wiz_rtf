# encoding: utf-8

# WizRft:  A gem for exporting Word Documents in ruby
# using the Microsoft Rich Text Format (RTF) Specification
# Copyright (C) 2015 by sgzhe@163.com

module WizRtf

  # = Represents an image
  # This class represents an image within a RTF document. Currently only the
  # PNG, JPEG, GIF and Windows Bitmap formats are supported.
  class Image
    JPEG_SOF_BLOCKS = [0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF]
    PIC_TYPE = {png: :pngblip, jpg: :jpegblip, bmp: :pngblip, gif: :pngblip}

    # This is the constructor for the Image class.
    # +file+ - image file path and filename.
    def initialize(file)
      begin
        @img = IO.binread(file)
        @width, @height = self.size
      rescue  Exception => e
        STDERR.puts "** error parsing #{file}: #{e.inspect}"
      end
    end

    # Returns an symbol indicating the image type fetched from a image file.
    # It will return nil if the image could not be fetched, or if the image type was not recognised.
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
                  nil
              end
    end

    # Returns an array containing the width and height of the image.
    # It will return nil if the image could not be fetched, or if the image type was not recognised.
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

    # Outputs the Partial Rtf Document to a Generic Stream as a Rich Text Format (RTF).
    # +io+ - The Generic IO to Output the RTF Document.
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