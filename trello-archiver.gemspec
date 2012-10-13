# -*- encoding: utf-8 -*-
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'trello-archiver/version'

Gem::Specification.new do |gem|
  gem.name          = "trello-archiver"
  gem.version       = TrelloArchiver::VERSION
  gem.authors       = ["MadTypist, ZPH"]
  gem.email         = ["MadTypist on Github, Zander@civet.ws"]
  gem.description   = %q{Trello Archiver library to export as XLSX or CSV}
  gem.summary       = %q{Provides two scripts in bin folder. trello_backup selectively backs up a single board, while trello_autoarchive backs up all boards.  The configuration is set in config.yml.}
  gem.homepage      = ""

  gem.files         = `git ls-files`.split($/)
  gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.test_files    = gem.files.grep(%r{^(test|spec|features)/})
  gem.require_paths = ["lib"]

  gem.add_dependency "ruby-trello"
  gem.add_dependency "xlsx_writer"

  gem.add_development_dependency "rake"
  gem.add_development_dependency "minitest"
  %w[guard guard-minitest growl guard-bundler rb-fsevent vcr fakeweb pry].each do |dep|
    gem.add_development_dependency dep
  end
end
