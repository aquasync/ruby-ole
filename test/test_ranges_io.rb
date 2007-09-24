#! /usr/bin/ruby

TEST_DIR = File.dirname __FILE__
$: << "#{TEST_DIR}/../lib"

require 'test/unit'
require 'ole/io_helpers'
require 'stringio'

class TestRangesIO < Test::Unit::TestCase
	def setup
		# read from ourself, also using overlaps.
		ranges = [100..200, 0..10, 100..150]
		@io = RangesIO.new open("#{TEST_DIR}/test_ranges_io.rb"), ranges, :close_parent => true
	end

	def teardown
		@io.close
	end

	def test_basics
		assert_equal 160, @io.size
		assert_match %r{size=160,.*range=100\.\.200}, @io.inspect
	end

	def test_truncate
		assert_raises(NotImplementedError) { @io.size += 10 }
	end

	def test_offset_and_size
		assert_equal [[100, 100], 0], @io.offset_and_size(0)
		assert_equal [[150, 50], 0], @io.offset_and_size(50)
		assert_equal [[5, 5], 1], @io.offset_and_size(105)
		assert_raises(ArgumentError) { @io.offset_and_size 1000 }
	end

	def test_seek
		@io.pos = 10
		@io.seek(-100, IO::SEEK_END)
		@io.seek(-10, IO::SEEK_CUR)
		@io.pos += 20
		assert_equal 70, @io.pos
		# seeking past the end doesn't throw an exception for normal
		# files, even in read mode, but RangesIO does
		assert_raises(Errno::EINVAL) { @io.seek 500 }
		assert_raises(Errno::EINVAL) { @io.seek(-500, IO::SEEK_END) }
		assert_raises(Errno::EINVAL) { @io.seek 1, 10 }
	end

	def test_read
		# this will map to the start of the file:
		@io.pos = 100
		assert_equal '#! /usr/bi', @io.read(10)
		# test selection of initial range, offset within that range
		pos = 80
		@io.seek pos
		# test advancing of pos properly, by...
		chunked = (0...10).map { @io.read 10 }.join
		# given the file is 160 long:
		assert_equal 80, chunked.length
		@io.seek pos
		# comparing with a flat read
		assert_equal chunked, @io.read(80)
	end

	# should test gets, lineno, and other IO methods we want to have
	def test_gets
		assert_equal "equire 'ole/io_helpers'\n", @io.gets
	end

	# would need to move to StringIO to do this.
	def test_write
	end
end

