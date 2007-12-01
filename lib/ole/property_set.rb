
module Ole
	module Types
		# should have a list of the VT_* variant types, and have all the serialization related code
		# here... implement dump & load functions like marshalling
		class Guid
			SIZE = 16

			def self.load str
				Types.load_guid str
			end
		end

		# see http://poi.apache.org/hpsf/internals.html
		class PropertySet
			HEADER_SIZE = 28
			HEADER_UNPACK = "vvVa#{Guid::SIZE}V"
			OS_MAP = {
				0 => :win16,
				1 => :mac,
				2 => :win32
			}

			class Section < Struct.new(:guid, :offset)
				include Enumerable

				SIZE = Guid::SIZE + 4
				UNPACK_STR = "a#{Guid::SIZE}v"

				attr_reader :length
				def initialize str, property_set
					@property_set = property_set
					super(*str.unpack(UNPACK_STR))
					self.guid = Guid.load guid
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
						type = io.read(4).unpack('V')[0]
						yield id, type, io.read(10)
					end
					self
				end

				def properties
					to_a
				end
			end

			attr_reader :io, :signature, :unknown, :os, :guid, :sections
			def initialize io
				@io = io
				load_header io.read(HEADER_SIZE)
				load_section_list io.read(@num_sections * Section::SIZE)
				# expect no gap between last section and start of data.
				Log.warn "gap between section list and property data" unless io.pos == @sections.map(&:offset).min
			end

			def load_header str
				@signature, @unknown, @os_id, @guid, @num_sections = str.unpack HEADER_UNPACK
				# should i check that unknown == 0? it usually is. so is the guid actually
				@guid = Guid.load @guid
				@os = OS_MAP[@os_id] || Log.warn("unknown operating system id #{@os_id}")
			end

			def load_section_list str
				@sections = str.scan(/.{#{Section::SIZE}}/m).map { |str| Section.new str, self }
			end
		end
	end
end

