# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'fast_xls_to_hash/version'

Gem::Specification.new do |spec|
  spec.name          = "fast_xls_to_hash"
  spec.version       = FastXlsToHash::VERSION
  spec.authors       = ["woodcrust"]
  spec.email         = ["roboucrop@gmail.com"]

  spec.summary       = %q{This is xls parser}
  spec.description   = %q{Fast parsing xls to the hash}
  spec.homepage      = "https://github.com/woodcrust/fast_xls_to_hash"
  spec.license       = "MIT"

  spec.files         = Dir['lib/**/*']
  spec.bindir        = "bin"
  spec.executables   = ["fast_xls_to_hash"]
  spec.require_paths = ["lib"]

  spec.test_files    = `git ls-files -- {test,spec,features}/*`.split("\n")

  spec.add_runtime_dependency     "spreadsheet", "~> 1.0"
  
  spec.add_development_dependency "bundler", "~> 1.12"
  spec.add_development_dependency "rake", "~> 10.0"
  spec.add_development_dependency "rspec", "~> 3.0"
  spec.add_development_dependency 'byebug'
  # spec.add_development_dependency "spreadsheet", "~> 1.0"
end
