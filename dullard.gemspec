# -*- encoding: utf-8 -*-
require File.expand_path('../lib/dullard/version', __FILE__)

Gem::Specification.new do |gem|
  gem.authors       = ["Ted Kaplan"]
  gem.email         = ["ted@shlashdot.org"]
  gem.summary       = %q{A fast XLSX parser using Nokogiri}
  gem.homepage      = "http://github.com/thirtyseven/dullard"
  gem.license       = "MIT"

  gem.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  gem.files         = `git ls-files`.split("\n")
  gem.test_files    = `git ls-files -- {test,spec,features}/*`.split("\n")
  gem.name          = "dullard"
  gem.require_paths = ["lib"]
  gem.version       = Dullard::VERSION
  gem.required_ruby_version = '~> 2.0'

  gem.add_development_dependency "rspec", "~> 2.14"
  gem.add_development_dependency "rake", "~> 10.1"
  gem.add_dependency "nokogiri", "~> 1.6"
  gem.add_dependency "rubyzip", "~> 1.1"
end
