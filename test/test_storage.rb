#! /usr/bin/ruby

$: << File.dirname(__FILE__) + '/../lib'

require 'test/unit'
require 'ole/storage'
require 'digest/sha1'
require 'stringio'

#
# = TODO
#
# These tests could be a lot more complete.
#

# should test resizeable and migrateable IO.

class TestStorageRead < Test::Unit::TestCase
	TEST_DIR = File.dirname __FILE__

	def setup
		@ole = Ole::Storage.open "#{TEST_DIR}/test_word_6.doc", 'rb'
	end

	def teardown
		@ole.close
	end

	def test_header
		# should have further header tests, testing the validation etc.
		assert_equal 17,  @ole.header.to_a.length
		assert_equal 117, @ole.header.dirent_start
		assert_equal 1,   @ole.header.num_bat
		assert_equal 1,   @ole.header.num_sbat
		assert_equal 0,   @ole.header.num_mbat
	end

	def test_fat
		# the fat block has all the numbers from 5..118 bar 117
		bbat_table = [112] + ((5..118).to_a - [112, 117])
		assert_equal bbat_table, @ole.bbat.reject { |i| i >= (1 << 32) - 3 }, 'bbat'
		sbat_table = (1..43).to_a - [2, 3]
		assert_equal sbat_table, @ole.sbat.reject { |i| i >= (1 << 32) - 3 }, 'sbat'
	end

	def test_directories
		assert_equal 5, @ole.dirents.length, 'have all directories'
		# a more complicated one would be good for this
		assert_equal 4, @ole.root.children.length, 'properly nested directories'
	end

	def test_utf16_conversion
		assert_equal 'Root Entry', @ole.root.name
		assert_equal 'WordDocument', @ole.root.children[2].name
	end

	def test_read
		# test the ole storage type
		type = 'Microsoft Word 6.0-Dokument'
		assert_equal type, (@ole.root/"\001CompObj").read[/^.{32}([^\x00]+)/m, 1]
		# i was actually not loading data correctly before, so carefully check everything here
		hashes = [-482597081, 285782478, 134862598, -863988921]
		assert_equal hashes, @ole.root.children.map { |child| child.read.hash }
	end
end

class TestStorageWrite < Test::Unit::TestCase
	TEST_DIR = File.dirname __FILE__

	def sha1 str
		Digest::SHA1.hexdigest str
	end

	# try and test all the various things the #flush function does
	def test_flush
	end
	
	# FIXME
	# don't really want to lock down the actual internal api's yet. this will just
	# ensure for the time being that #flush continues to work properly. need a host
	# of checks involving writes that resize their file bigger/smaller, that resize
	# the bats to more blocks, that resizes the sb_blocks, that has migration etc.
	def test_write_hash
		io = StringIO.open File.read("#{TEST_DIR}/test_word_6.doc")
		assert_equal '9974e354def8471225f548f82b8d81c701221af7', sha1(io.string)
		Ole::Storage.open(io) { }
		assert_equal 'efa8cfaf833b30b1d1d9381771ddaafdfc95305c', sha1(io.string)
		# add a repack test here
		Ole::Storage.open io, &:repack
		assert_equal 'c8bb9ccacf0aaad33677e1b2a661ee6e66a48b5a', sha1(io.string)
	end

	def test_plain_repack
		io = StringIO.open File.read("#{TEST_DIR}/test_word_6.doc")
		assert_equal '9974e354def8471225f548f82b8d81c701221af7', sha1(io.string)
		Ole::Storage.open io, &:repack
		# note equivalence to the above flush, repack, flush
		assert_equal 'c8bb9ccacf0aaad33677e1b2a661ee6e66a48b5a', sha1(io.string)
	end

	def test_create_from_scratch_hash
		io = StringIO.new
		Ole::Storage.open(io) { }
		assert_equal '6bb9d6c1cdf1656375e30991948d70c5fff63d57', sha1(io.string)
		# more repack test, note invariance
		Ole::Storage.open io, &:repack
		assert_equal '6bb9d6c1cdf1656375e30991948d70c5fff63d57', sha1(io.string)
	end

	def test_create_dirent
		Ole::Storage.open StringIO.new do |ole|
			dirent = Ole::Storage::Dirent.new ole, :name => 'test name', :type => :dir
			assert_equal 'test name', dirent.name
			assert_equal :dir, dirent.type
			# for a dirent created from scratch, type_id is currently not set until serialization:
			assert_equal 0, dirent.type_id
			dirent.to_s
			assert_equal 1, dirent.type_id
			assert_raises(ArgumentError) { Ole::Storage::Dirent.new ole, :type => :bogus }
		end
	end
end

