#! /usr/bin/ruby

$: << File.dirname(__FILE__) + '/../lib'

require 'test/unit'
require 'ole/storage'
require 'digest/sha1'
require 'stringio'
require 'tempfile'
require 'zlib'
require 'base64'

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
	
	def test_new_without_explicit_mode
		open "#{TEST_DIR}/test_word_6.doc", 'rb' do |f|
			assert_equal false, Ole::Storage.new(f).writeable
		end
	end

	def capture_warnings
		@warn = []
		outer_warn = @warn
		old_log = Ole::Log
		old_verbose = $VERBOSE
		begin
			$VERBOSE = nil
			Ole.const_set :Log, Object.new
			# restore for the yield
			$VERBOSE = old_verbose
			(class << Ole::Log; self; end).send :define_method, :warn do |message|
				outer_warn << message
			end
			yield
		ensure
			$VERBOSE = nil
			Ole.const_set :Log, old_log
			$VERBOSE = old_verbose
		end
	end

	def test_invalid
		assert_raises Ole::Storage::FormatError do
			Ole::Storage.open StringIO.new(0.chr * 1024)
		end
		assert_raises Ole::Storage::FormatError do
			Ole::Storage.open StringIO.new(Ole::Storage::Header::MAGIC + 0.chr * 1024)
		end
		capture_warnings do
			head = Ole::Storage::Header.new
			head.threshold = 1024
			assert_raises NoMethodError do
				Ole::Storage.open StringIO.new(head.to_s + 0.chr * 1024)
			end
		end
		assert_equal ['may not be a valid OLE2 structured storage file'], @warn
	end
	
	def test_inspect
		assert_match(/#<Ole::Storage io=#<File:.*?test_word_6.doc> root=#<Dirent:"Root Entry">>/, @ole.inspect)
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
		# the regular String#hash was different on the mac, so asserting
		# against full strings. switch this to sha1 instead of this fugly blob
		data = <<-end
			eJxjZGBgYkADAABKAAQ=

			eJxjZPj3n5mLgeE/EDBwMjGAwAEwyeAmAyR8M5OL8ovz00oUwvOLUhTM9Ax0
			XfKzS3NT80oYuEDywSBxl/xkBgEgD8TWA3LA8npmYGMAKWQZqA==

			eJztVs9rE0EU/mZ32m5smyZtUlKQunhIezAV6knwYGoRxRLQBG89zCabdiE7
			K9kJpt7Es1DwLxCsP+jJqycv/S/8I9SrmPgmuy0pK1ilNZd8MPt233zz7dud
			H+99OXBR+pziILgYgglkqaUvIYFs0pWEBSyzIfsHLKeA3Fl0Y6wTX4d2K7bH
			uEvP6i90zhuf6P21fxpp0C9nOMOvGuMcUQ1811ZuqOjSVWuzszm8eb52cFR+
			K/k7yd9L/kHyV6OO8gJBezwT7/UVaj/xY9QRjXHhMDoM5rapZ39o/ntRZ2+0
			sY3xv0C5xkhT1rXo7gHmBr5VWg7f+gbZqU23KTotSqaT5L+Ba2waDmjR1AsQ
			qZnf6BVRvv29/5rsUtkJhXpWqiohG6LdCOu7ba+pRFud5veuMBisiKl7rmh4
			ckfn8jnkv2KxC4tNYpujfhk2NjKaZyNVo0PadoLGni4sMshDM3XlcD3D2DzL
			gW95oaJcmj3Rv6r174gnyguk1p9HvqtHpdmE/pTHDIUBb50VMHFfNlwS5Fig
			/ihKrcQH+wPoExCflMZJjWRSxZFf3Cd+zfPd0K64T+1HgS8kZshroLrnO0EL
			00VNKbc90cKqpi9UPN/phBHXrgQ37a2EwpIelI6JVSFD4kQSGxQVs14WgCMK
			ZXcQ7WHtYxPWw5JupyuJqLLgeLHPEt5jz4r5CxXWcbY=

			eJzt0z1oU2EcRvE3rR+tWtGtOBQ3hS4tODgWKnQRlM5dHAoZCuIHaLcsBaGT
			hIBksQmlQ1wCTgF3Q+Z0cHN1lZBQQhL/N60fUEEKh5bK+cHleXlJw+2hHY5S
			yqXjsruvW++/HzzK3/jwdirN3/n4ZSHu2vspXY+9Gc/ro88txN107P3YK7EP
			Yq/GPo6diV2LvRabj52KfRl7IbYQeyn2Xezl2N3Yi7H12MnYT7ETsc3Yley+
			ndIoZN9xN55n8eyM32F2/M6F9vHfYyTa3lm/wH/Ipjyb8mzKsynPpjyb8mzK
			synPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzK
			synPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzK
			synPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzK
			synPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzK
			synPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzK
			synPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzK
			synPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzKsynPpjyb8mzKsynPprx/NO33
			+7VaLQ6VSqVUKpXL5cFg0Gq1Go1GtVptNpvFYrFer8dNp9M5lRc+B078dzoc
			Dnu93uGh2+3G+deNDvm/z7Mpz6Y8m/JsyrMpz6Y8m/JsyrMpz6a8vYmU0rdB
			SnOx2XkxnqX02/j8ZnXzb+fCrcr09r3Puexnc0efz84z8SznnzzfWH9x++HT
			V+s/7//8zEnOPwDhN6kw
		end
		expect = data.split(/\n\s*\n/).map { |chunk| Zlib::Inflate.inflate Base64.decode64(chunk) }

		# test the ole storage type
		type = 'Microsoft Word 6.0-Dokument'
		assert_equal type, (@ole.root/"\001CompObj").read[32..-1][/([^\x00]+)/m, 1]
		# i was actually not loading data correctly before, so carefully check everything here
		assert_equal expect, @ole.root.children.map { |child| child.read }
	end

	def test_dirent
		dirent = @ole.root.children.first
		assert_equal "\001Ole", dirent.name
		assert_equal 20, dirent.size
		assert_equal '#<Dirent:"Root Entry">', @ole.root.inspect
		
		# exercise Dirent#[]. note that if you use a number, you get the Struct
		# fields.
		assert_equal dirent, @ole.root["\001Ole"]
		assert_equal dirent.name_utf16, dirent[0]
		assert_equal nil, @ole.root.time
		
		assert_equal @ole.root.children, @ole.root.to_enum(:each_child).to_a

		dirent.open('r') { |f| assert_equal 2, f.first_block }
		dirent.open('w') { |f| }
		assert_raises Errno::EINVAL do
			dirent.open('a') { |f| }
		end
	end

	def test_delete
		dirent = @ole.root.children.first
		assert_raises(ArgumentError) { @ole.root.delete nil }
		assert_equal [dirent], @ole.root.children & [dirent]
		assert_equal 20, dirent.size
		@ole.root.delete dirent
		assert_equal [], @ole.root.children & [dirent]
		assert_equal 0, dirent.size
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
		io = StringIO.open open("#{TEST_DIR}/test_word_6.doc", 'rb', &:read)
		assert_equal '9974e354def8471225f548f82b8d81c701221af7', sha1(io.string)
		Ole::Storage.open(io, :update_timestamps => false) { }
		# hash changed. used to be efa8cfaf833b30b1d1d9381771ddaafdfc95305c
		# thats because i know truncate the io, and am probably removing some trailing allocated
		# available blocks.
		assert_equal 'a39e3c4041b8a893c753d50793af8d21ca8f0a86', sha1(io.string)
		# add a repack test here
		Ole::Storage.open io, :update_timestamps => false, &:repack
		assert_equal 'c8bb9ccacf0aaad33677e1b2a661ee6e66a48b5a', sha1(io.string)
	end

	def test_plain_repack
		io = StringIO.open open("#{TEST_DIR}/test_word_6.doc", 'rb', &:read)
		assert_equal '9974e354def8471225f548f82b8d81c701221af7', sha1(io.string)
		Ole::Storage.open io, :update_timestamps => false, &:repack
		# note equivalence to the above flush, repack, flush
		assert_equal 'c8bb9ccacf0aaad33677e1b2a661ee6e66a48b5a', sha1(io.string)
		# lets do it again using memory backing
		Ole::Storage.open(io, :update_timestamps => false) { |ole| ole.repack :mem }
		# note equivalence to the above flush, repack, flush
		assert_equal 'c8bb9ccacf0aaad33677e1b2a661ee6e66a48b5a', sha1(io.string)
		assert_raises ArgumentError do
			Ole::Storage.open(io, :update_timestamps => false) { |ole| ole.repack :typo }
		end
	end

	def test_create_from_scratch_hash
		io = StringIO.new('')
		Ole::Storage.open(io) { }
		assert_equal '6bb9d6c1cdf1656375e30991948d70c5fff63d57', sha1(io.string)
		# more repack test, note invariance
		Ole::Storage.open io, :update_timestamps => false, &:repack
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

