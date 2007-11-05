#! /usr/bin/ruby

$: << File.dirname(__FILE__) + '/../lib'

require 'test/unit'
require 'ole/support'

class TestSupport < Test::Unit::TestCase
	TEST_DIR = File.dirname __FILE__

	def test_file
		assert_equal 4096, open("#{TEST_DIR}/oleWithDirs.ole") { |f| f.size }
		# point is to have same interface as:
		assert_equal 4096, StringIO.open(File.read("#{TEST_DIR}/oleWithDirs.ole")).size
	end

	def test_enumerable
		expect = {0 => [2, 4], 1 => [1, 3]}
		assert_equal expect, [1, 2, 3, 4].group_by { |i| i & 1 }
		assert_equal 10, [1, 2, 3, 4].sum
		assert_equal %w[1 2 3 4], [1, 2, 3, 4].map(&:to_s)
	end

	def test_logger
		io = StringIO.new
		log = Logger.new_with_callstack io
		log.warn 'test'
		expect = %r{^\[\d\d:\d\d:\d\d \./test/test_support\.rb:\d+:test_logger\]\nWARN   test$}
		assert_match expect, io.string.chomp
	end

	def test_io
		str = 'a' * 5000 + 'b'
		src, dst = StringIO.new(str), StringIO.new
		IO.copy src, dst
		assert_equal str, dst.string
	end

	def test_string
		str = "aa aa ||| aa aa"
		assert_equal [0, 3, 10, 13], str.indexes('aa')
		# this is mostly a check that regexp quote is used.
		assert_equal [6, 7, 8], str.indexes('|')
		# note not [6, 7] - no overlaps
		assert_equal [6], str.indexes('||')
	end
end

class TestRecursivelyEnumerable < Test::Unit::TestCase
	class Container
		include RecursivelyEnumerable
	
		def initialize *children
			@children = children
		end
	
		def each_child(&block)
			@children.each(&block)
		end
	
		def inspect
			"#<Container>"
		end
	end
	
	def setup
		@root = Container.new(
			Container.new(1),
			Container.new(2,
				Container.new(
					Container.new(3)
				)
			),
			4,
			Container.new()
		)
	end

	def test_find
		i = 0
		found = @root.recursive.find do |obj|
			i += 1
			obj == 4
		end
		assert_equal found, 4
		assert_equal 9, i

		i = 0
		found = @root.recursive(:breadth_first).find do |obj|
			i += 1
			obj == 4
		end
		assert_equal found, 4
		assert_equal 4, i

		# this is to make sure we hit the breadth first child cache
		i = 0
		found = @root.recursive(:breadth_first).find do |obj|
			i += 1
			obj == 3
		end
		assert_equal found, 3
		assert_equal 10, i
	end

	def test_to_tree
		assert_equal <<-'end', @root.to_tree
- #<Container>
  |- #<Container>
  |  \- 1
  |- #<Container>
  |  |- 2
  |  \- #<Container>
  |     \- #<Container>
  |        \- 3
  |- 4
  \- #<Container>
		end
	end
end

