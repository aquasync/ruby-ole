#! /usr/bin/ruby

$: << File.dirname(__FILE__) + '/../lib'

require 'test/unit'
require 'ole/support'
require 'stringio'

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
end

