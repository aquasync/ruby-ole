# encoding: ASCII-8BIT

module Ole
	module Types
		#
		# The PropertySet class currently supports readonly access to the properties
		# serialized in "property set" streams, such as the file "\005SummaryInformation",
		# in OLE files.
		#
		# Think it has its roots in MFC property set serialization.
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
			DATA = {
				Clsid.parse('{f29f85e0-4ff9-1068-ab91-08002b27b3d9}') => ['FMTID_SummaryInformation', {
					2  => 'doc_title',
					3  => 'doc_subject',
					4  => 'doc_author',
					5  => 'doc_keywords',
					6  => 'doc_comments',
					7  => 'doc_template',
					8  => 'doc_last_author',
					9  => 'doc_rev_number',
					10 => 'doc_edit_time',
					11 => 'doc_last_printed',
					12 => 'doc_created_time',
					13 => 'doc_last_saved_time',
					14 => 'doc_page_count',
					15 => 'doc_word_count',
					16 => 'doc_char_count',
					18 => 'doc_app_name',
					19 => 'security'
				}],
				Clsid.parse('{d5cdd502-2e9c-101b-9397-08002b2cf9ae}') => ['FMTID_DocSummaryInfo', {
					2  => 'doc_category',
					3  => 'doc_presentation_target',
					4  => 'doc_byte_count',
					5  => 'doc_line_count',
					6  => 'doc_para_count',
					7  => 'doc_slide_count',
					8  => 'doc_note_count',
					9  => 'doc_hidden_count',
					10 => 'mmclips',
					11 => 'scale_crop',
					12 => 'heading_pairs',
					13 => 'doc_part_titles',
					14 => 'doc_manager',
					15 => 'doc_company',
					16 => 'links_up_to_date'
				}],
				Clsid.parse('{d5cdd505-2e9c-101b-9397-08002b2cf9ae}') => ['FMTID_UserDefinedProperties', {}]
			}

			# create an inverted map of names to guid/key pairs
			PROPERTY_MAP = DATA.inject({}) do |h1, (guid, data)|
				data[1].inject(h1) { |h2, (id, name)| h2.update name => [guid, id] }
			end

			module Constants
				DATA.each { |guid, (name, _)| const_set name, guid }
			end

			include Constants
			include Enumerable

			class Section
				include Variant::Constants
				include Enumerable

				SIZE = Clsid::SIZE + 4
				PACK = "a#{Clsid::SIZE}v"

				attr_accessor :guid, :offset
				attr_reader :length

				def initialize str, property_set
					@property_set = property_set
					@guid, @offset = str.unpack PACK
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
				
				def [] key
					each_raw do |id, property_offset|
						return read_property(property_offset).last if key == id
					end
					nil
				end
				
				def []= key, value
					raise NotImplementedError, 'section writes not yet implemented'
				end
				
				def each
					each_raw do |id, property_offset|
						yield id, read_property(property_offset).last
					end
				end

			private

				def each_raw
					io.seek offset + 8
					io.read(length * 8).each_chunk(8) { |str| yield(*str.unpack('V2')) }
				end
				
				def read_property property_offset
					io.seek offset + property_offset
					type, value = io.read(8).unpack('V2')
					# is the method of serialization here custom?
					case type
					when VT_LPSTR, VT_LPWSTR
						value = Variant.load type, io.read(value)
					# ....
					end
					[type, value]
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
				@signature, @unknown, @os_id, @guid, @num_sections = str.unpack HEADER_PACK
				# should i check that unknown == 0? it usually is. so is the guid actually
				@guid = Clsid.load @guid
				@os = OS_MAP[@os_id] || Log.warn("unknown operating system id #{@os_id}")
			end

			def load_section_list str
				@sections = str.to_enum(:each_chunk, Section::SIZE).map { |s| Section.new s, self }
			end
			
			def [] key
				pair = PROPERTY_MAP[key.to_s] or return nil
				section = @sections.find { |s| s.guid == pair.first } or return nil
				section[pair.last]
			end
			
			def []= key, value
				pair = PROPERTY_MAP[key.to_s] or return nil
				section = @sections.find { |s| s.guid == pair.first } or return nil
				section[pair.last] = value
			end
			
			def method_missing name, *args, &block
				if name.to_s =~ /(.*)=$/
					return super unless args.length == 1
					return super unless PROPERTY_MAP[$1]
					self[$1] = args.first
				else
					return super unless args.length == 0
					return super unless PROPERTY_MAP[name.to_s]
					self[name]
				end
			end
			
			def each
				@sections.each do |section|
					next unless pair = DATA[section.guid]
					map = pair.last
					section.each do |id, value|
						name = map[id] or next
						yield name, value
					end
				end
			end
			
			def to_h
				inject({}) { |hash, (name, value)| hash.update name.to_sym => value }
			end
		end
	end
end

