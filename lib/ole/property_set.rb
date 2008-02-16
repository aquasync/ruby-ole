require 'ole/types'
require 'yaml'

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
			HEADER_PACK = "vvVa#{Clsid::SIZE}V"
			OS_MAP = {
				0 => :win16,
				1 => :mac,
				2 => :win32,
				0x20001 => :ooffice, # open office on linux...
			}

			# define a smattering of the property set guids. 
			#FMTID_SummaryInformation		= Clsid.parse '{f29f85e0-4ff9-1068-ab91-08002b27b3d9}'
			#FMTID_DocSummaryInformation	= Clsid.parse '{d5cdd502-2e9c-101b-9397-08002b2cf9ae}'
			#FMTID_UserDefinedProperties	= Clsid.parse '{d5cdd505-2e9c-101b-9397-08002b2cf9ae}'

			DATA = YAML.load_file(File.dirname(__FILE__) + '/../../data/propids.yaml').
				inject({}) { |hash, (key, value)| hash.update Clsid.parse(key) => value }

			module Constants
				DATA.each { |guid, (name, map)| const_set name, guid }
			end

			include Constants

			class Section < Struct.new(:guid, :offset)
				include Variant::Constants
				include Enumerable

				SIZE = Clsid::SIZE + 4
				PACK = "a#{Clsid::SIZE}v"

				attr_reader :length
				def initialize str, property_set
					@property_set = property_set
					super(*str.unpack(PACK))
					self.guid = Clsid.load guid
					@map = DATA[guid] ? DATA[guid][1] : nil
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

				def [] key
					unless Integer === key
						return unless @map and key = @map.invert[key]
					end
					return unless result = properties.assoc(key)
					result.last
				end

				def method_missing name, *args
					if args.empty? and @map and @map.values.include? name.to_s
						self[name.to_s]
					else
						super
					end
				end

				def properties
					@properties ||= to_enum.to_a
				end

				#def to_h
				#	properties.inject({}) do |hash, (key, type, value)|
				#		hash.update 
				#end
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
				@signature, @unknown, @os_id, @guid, @num_sections = str.unpack HEADER_PACK
				# should i check that unknown == 0? it usually is. so is the guid actually
				@guid = Clsid.load @guid
				@os = OS_MAP[@os_id] || Log.warn("unknown operating system id #{@os_id}")
			end

			def load_section_list str
				@sections = str.scan(/.{#{Section::SIZE}}/m).map { |s| Section.new s, self }
			end
		end
	end
	
	class Storage
		# i'm thinking - search for a property set in +filenames+ containing a
		# section with guid +guid+. then yield it. can read/write to it in the
		# block.
		# propsets themselves can have guids, but they are often all null.
		def with_property_set guid, filenames=nil
		end

		class PropertySetSectionProxy
			attr_reader :obj, :section_num
			def initialize obj, section_num
				@obj, @section_num = obj, section_num
			end
			
			def method_missing name, *args, &block
				obj.open do |io|
					section = Types::PropertySet.new(io).sections[section_num]
					section.send name, *args, &block
				end
			end
		end

		# this will be changed to use with_property_set
		def summary_information
			dirent = root["\005SummaryInformation"]
			dirent.open do |io|
				propset = Types::PropertySet.new(io)
				sections = propset.sections
				# this will maybe get wrapped up as
				# section = propset[guid]
				# maybe taking it one step further, i'd hide the section thing,
				# and let you use composite keys, like
				# propset[4, guid] eg in MAPI, and just propset.doc_author.
				section = sections.find do |s|
					s.guid == Types::PropertySet::FMTID_SummaryInformation
				end
				return PropertySetSectionProxy.new(dirent, sections.index(section))
			end
		end
		
		alias summary_info :summary_information
	end
end

