# WizRtf

A gem for exporting Word Documents in ruby using the Microsoft Rich Text Format (RTF) Specification.

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'wiz_rtf'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install wiz_rtf

## Usage
```ruby
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
```
## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release` to create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

1. Fork it ( https://github.com/[my-github-username]/wiz_rtf/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request
