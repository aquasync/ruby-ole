#! /usr/bin/ruby

$: << File.dirname(__FILE__) + '/../lib'

require 'test/unit'
require 'ole/types/property_set'
require 'ole/storage/base'
require 'ole/storage/meta_data'

class TestTypes < Test::Unit::TestCase
	include Ole::Types

	def setup
		@io = open File.dirname(__FILE__) + '/test_SummaryInformation', 'rb'
	end

	def teardown
		@io.close
	end

	def test_property_set
		propset = PropertySet.new @io
		assert_equal :mac, propset.os
		assert_equal 1, propset.sections.length
		section = propset.sections.first
		assert_equal 14, section.length
		assert_equal 'f29f85e0-4ff9-1068-ab91-08002b27b3d9', section.guid.format
		assert_equal PropertySet::FMTID_SummaryInformation, section.guid
		assert_equal 'Charles Lowe', section.to_a.assoc(4).last
		# new named support
		assert_equal 'Charles Lowe', propset.doc_author
	end

	# should be expanded, and moved to test_meta_data.rb
	def test_ole_storage_integration
		Ole::Storage.open File.dirname(__FILE__) + '/test.doc', 'rb' do |ole|
			assert_equal 'Charles Lowe', ole.meta_data.doc_author
			assert_equal 'Title', ole.meta_data.doc_title
		end
	end
end

