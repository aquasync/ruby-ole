require 'ole/types'

module Ole
	module Types
		#
		# The PropertySet class currently supports readonly access to the properties
		# serialized in "property set" streams, such as the file "\005SummaryInformation",
		# in OLE files.
		#
		# Think has its roots in MFC property set serialization.
		#
		# See http://poi.apache.org/hpsf/internals.html for details
		#
		class PropertySet
			HEADER_SIZE = 28
			HEADER_UNPACK = "vvVa#{Clsid::SIZE}V"
			OS_MAP = {
				0 => :win16,
				1 => :mac,
				2 => :win32
			}

			# define a smattering of the property set guids. 
			FMTID_SummaryInformation		= Clsid.parse '{f29f85e0-4ff9-1068-ab91-08002b27b3d9}'
			FMTID_DocSummaryInformation	= Clsid.parse '{d5cdd502-2e9c-101b-9397-08002b2cf9ae}'
			FMTID_UserDefinedProperties	= Clsid.parse '{d5cdd505-2e9c-101b-9397-08002b2cf9ae}'

			class Section < Struct.new(:guid, :offset)
				include Variant::Constants
				include Enumerable

				SIZE = Clsid::SIZE + 4
				UNPACK_STR = "a#{Clsid::SIZE}v"

				attr_reader :length
				def initialize str, property_set
					@property_set = property_set
					super(*str.unpack(UNPACK_STR))
					self.guid = Clsid.load guid
					load_header
				end

				def io
					@property_set.io
				end

				def load_header
					io.seek offset
					@byte_size, @length = io.read(8).unpack 'V2'
				end

				def each
					io.seek offset + 8
					io.read(length * 8).scan(/.{8}/m).each do |str|
						id, property_offset = str.unpack 'V2'
						io.seek offset + property_offset
						type, value = io.read(8).unpack('V2')
						# is the method of serialization here custom?
						case type
						when VT_LPSTR, VT_LPWSTR
							value = Variant.load type, io.read(value)
						# ....
						end
						yield id, type, value
					end
					self
				end

				def properties
					to_enum.to_a
				end
			end

			attr_reader :io, :signature, :unknown, :os, :guid, :sections
			def initialize io
				@io = io
				load_header io.read(HEADER_SIZE)
				load_section_list io.read(@num_sections * Section::SIZE)
				# expect no gap between last section and start of data.
				#Log.warn "gap between section list and property data" unless io.pos == @sections.map(&:offset).min
			end

			def load_header str
				@signature, @unknown, @os_id, @guid, @num_sections = str.unpack HEADER_UNPACK
				# should i check that unknown == 0? it usually is. so is the guid actually
				@guid = Clsid.load @guid
				@os = OS_MAP[@os_id] || Log.warn("unknown operating system id #{@os_id}")
			end

			def load_section_list str
				@sections = str.scan(/.{#{Section::SIZE}}/m).map { |str| Section.new str, self }
			end
		end
	end
end

