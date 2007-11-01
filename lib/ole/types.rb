require 'iconv'
require 'date'

require 'ole/base'

module Ole # :nodoc:
	# FIXME
	module Types
		#
		# The OLE variant types, extracted from
		# http://www.marin.clara.net/COM/variant_type_definitions.htm.
		#
		# A subset is also in WIN32OLE::VARIANT, but its not cross platform (obviously).
		#
		# Use like:
		#
		#   p Ole::Types::Variant::NAMES[0x001f] => 'VT_LPWSTR'
		#   p Ole::Types::VT_DATE # => 7
		#
		# The serialization / deserialization functions should be fixed to make it easier
		# to work with. like
		#
		#   Ole::Types.from_str(VT_DATE, data) # and
		#   Ole::Types.to_str(VT_DATE, data)
		#
		# Or similar, rather than having to do VT_* <=> ad hoc class name etc as it is
		# currently.
		#
		module Variant
			NAMES = {
				0x0000 => 'VT_EMPTY',
				0x0001 => 'VT_NULL',
				0x0002 => 'VT_I2',
				0x0003 => 'VT_I4',
				0x0004 => 'VT_R4',
				0x0005 => 'VT_R8',
				0x0006 => 'VT_CY',
				0x0007 => 'VT_DATE',
				0x0008 => 'VT_BSTR',
				0x0009 => 'VT_DISPATCH',
				0x000a => 'VT_ERROR',
				0x000b => 'VT_BOOL',
				0x000c => 'VT_VARIANT',
				0x000d => 'VT_UNKNOWN',
				0x000e => 'VT_DECIMAL',
				0x0010 => 'VT_I1',
				0x0011 => 'VT_UI1',
				0x0012 => 'VT_UI2',
				0x0013 => 'VT_UI4',
				0x0014 => 'VT_I8',
				0x0015 => 'VT_UI8',
				0x0016 => 'VT_INT',
				0x0017 => 'VT_UINT',
				0x0018 => 'VT_VOID',
				0x0019 => 'VT_HRESULT',
				0x001a => 'VT_PTR',
				0x001b => 'VT_SAFEARRAY',
				0x001c => 'VT_CARRAY',
				0x001d => 'VT_USERDEFINED',
				0x001e => 'VT_LPSTR',
				0x001f => 'VT_LPWSTR',
				0x0040 => 'VT_FILETIME',
				0x0041 => 'VT_BLOB',
				0x0042 => 'VT_STREAM',
				0x0043 => 'VT_STORAGE',
				0x0044 => 'VT_STREAMED_OBJECT',
				0x0045 => 'VT_STORED_OBJECT',
				0x0046 => 'VT_BLOB_OBJECT',
				0x0047 => 'VT_CF',
				0x0048 => 'VT_CLSID',
				0x0fff => 'VT_ILLEGALMASKED',
				0x0fff => 'VT_TYPEMASK',
				0x1000 => 'VT_VECTOR',
				0x2000 => 'VT_ARRAY',
				0x4000 => 'VT_BYREF',
				0x8000 => 'VT_RESERVED',
				0xffff => 'VT_ILLEGAL'
			}

			module Constants
				NAMES.each { |num, name| const_set name, num }
			end
		end

		include Variant::Constants

		# the rest of this file is all a bit of adhoc marshal/unmarshal stuff

		# for VT_LPWSTR
		FROM_UTF16 = Iconv.new 'utf-8', 'utf-16le'
		TO_UTF16   = Iconv.new 'utf-16le', 'utf-8'

		# for VT_FILETIME
		EPOCH = DateTime.parse '1601-01-01'
		# Create a +DateTime+ object from a struct +FILETIME+
		# (http://msdn2.microsoft.com/en-us/library/ms724284.aspx).
		#
		# Converts +str+ to two 32 bit time values, comprising the high and low 32 bits of
		# the 100's of nanoseconds since 1st january 1601 (Epoch).
		def self.load_time str
			low, high = str.unpack 'L2'
			# we ignore these, without even warning about it
			return nil if low == 0 and high == 0
			time = EPOCH + (high * (1 << 32) + low) / 1e7 / 86400 rescue return
			# extra sanity check...
			unless (1800...2100) === time.year
				Log.warn "ignoring unlikely time value #{time.to_s}"
				return nil
			end
			time
		end

		# +time+ should be able to be either a Time, Date, or DateTime.
		def self.save_time time
			# i think i'll convert whatever i get to be a datetime, because of
			# the covered range.
			return 0.chr * 8 unless time
			time = time.send(:to_datetime) if Time === time
			bignum = ((time - Ole::Types::EPOCH) * 86400 * 1e7.to_i)
			high, low = bignum.divmod 1 << 32
			[low, high].pack 'L2'
		end

		# for VT_CLSID
		# Convert a binary guid into a plain string (will move to proper class later).
		def self.load_guid str
			"{%08x-%04x-%04x-%02x%02x-#{'%02x' * 6}}" % str.unpack('L S S CC C6')
		end
	end
end

