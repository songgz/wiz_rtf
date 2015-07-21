# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'wiz_rtf/version'

Gem::Specification.new do |spec|
  spec.name          = "wiz_rtf"
  spec.version       = WizRtf::VERSION
  spec.authors       = ["songgz"]
  spec.email         = ["sgzhe@163.com"]

  spec.summary       = %q{A gem for exporting Word Documents in ruby using the Microsoft Rich Text Format (RTF) Specification.}
  spec.description   = %q{A gem for rtf.}
  spec.homepage      = "https://github.com/songgz/wiz_rtf"
  spec.license       = "MIT"

  # Prevent pushing this gem to RubyGems.org by setting 'allowed_push_host', or
  # delete this section to allow pushing this gem to any host.
  if spec.respond_to?(:metadata)
    spec.metadata['allowed_push_host'] = "https://rubygems.org"
  else
    raise "RubyGems 2.0 or newer is required to protect against public gem pushes."
  end

  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.9"
  spec.add_development_dependency "rake", "~> 10.0"
end
