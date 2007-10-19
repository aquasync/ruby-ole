require 'rake/rdoctask'
require 'rake/testtask'
require 'rake/packagetask'
require 'rake/gempackagetask'

require 'rbconfig'
require 'fileutils'

$:.unshift 'lib'

require 'ole/storage'

PKG_NAME = 'ruby-ole'
PKG_VERSION = Ole::Storage::VERSION

task :default => [:test]

Rake::TestTask.new(:test) do |t|
	t.test_files = FileList["test/test_*.rb"]
	t.warning = true
	t.verbose = true
end

# RDocTask wasn't working for me
desc 'Build the rdoc HTML Files'
task :rdoc do
	system "rdoc -S -N --main 'Ole::Storage' --tab-width 2 --title '#{PKG_NAME} documentation' lib"
end

spec = Gem::Specification.new do |s|
	s.name = PKG_NAME
	s.version = PKG_VERSION
	s.summary = %q{Ruby OLE library.}
	s.description = %q{A library for easy read/write access to OLE compound documents for Ruby.}
	s.authors = ["Charles Lowe"]
	s.email = %q{aquasync@gmail.com}
	s.homepage = %q{http://code.google.com/p/ruby-ole}
	#s.rubyforge_project = %q{ruby-ole}

	s.executables = ['oletool']
	s.files  = ['Rakefile']
	s.files += Dir.glob("lib/**/*.rb")
	s.files += Dir.glob("test/test_*.rb") + Dir.glob("test/*.doc")
	s.files += Dir.glob("bin/*")

	s.has_rdoc = true
	s.rdoc_options += ['--main', 'Ole::Storage',
					   '--title', "#{PKG_NAME} documentation",
					   '--tab-width', '2']

	s.autorequire = 'ole/storage'
end

Rake::GemPackageTask.new(spec) do |p|
	p.gem_spec = spec
	p.need_tar = false
	p.need_zip = false
	p.package_dir = 'build'
end

