#! /usr/bin/ruby

$: << File.dirname(__FILE__) + '/../lib'

require 'test/unit'
require 'ole/property_set'

class TestTypes < Test::Unit::TestCase
	include Ole::Types

	def setup
		@io = open File.dirname(__FILE__) + '/test_SummaryInformation'
	end

	def teardown
		@io.close
	end

	def test_property_set
		propset = PropertySet.new @io
		assert_equal 1, propset.sections.length
		section = propset.sections.first
		assert_equal 14, section.length
		assert_equal 'f29f85e0-4ff9-1068-ab91-08002b27b3d9', section.guid.format
		assert_equal PropertySet::FMTID_SummaryInformation, section.guid
		# i expect this null byte should have be stripped. need to fix the encoding functions.
		assert_equal "Charles Lowe\000", section.properties.assoc(4).last
	end
end

