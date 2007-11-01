
# 
# A file with general support functions used by most files in the project.
# 
# These are the only methods added to other classes.
# 

require 'logger'
require 'stringio'
require 'enumerator'

class String # :nodoc:
	# plural of String#index. returns all offsets of +string+. rename to indices?
	#
	# note that it doesn't check for overlapping values.
	def indexes string
		# in some ways i'm surprised that $~ works properly in this case...
		to_enum(:scan, /#{Regexp.quote string}/m).map { $~.begin 0 }
	end
end

class File # :nodoc:
	# for interface consistency with StringIO etc (rather than adding #stat
	# to them). used by RangesIO.
	def size
		stat.size
	end
end

class Symbol # :nodoc:
	def to_proc
		proc { |a| a.send self }
	end
end

module Enumerable # :nodoc:
	# 1.9 backport
	def group_by
		hash = Hash.new { |hash, key| hash[key] = [] }
		each { |item| hash[yield(item)] << item }
		hash
	end

	def sum initial=0
		inject(initial) { |a, b| a + b }
	end
end

# move to support?
class IO # :nodoc:
	def self.copy src, dst
		until src.eof?
			buf = src.read(4096)
			dst.write buf
		end
	end
end

class Logger # :nodoc:
	# A helper method for creating a +Logger+ which produce call stack
	# in their output
	def self.new_with_callstack logdev=STDERR
		log = Logger.new logdev
		log.level = WARN
		log.formatter = proc do |severity, time, progname, msg|
			# find where we were called from, in our code
			callstack = caller.dup
			callstack.shift while callstack.first =~ /\/logger\.rb:\d+:in/
			from = callstack.first.sub(/:in `(.*?)'/, ":\\1")
			"[%s %s]\n%-7s%s\n" % [time.strftime('%H:%M:%S'), from, severity, msg.to_s]
		end
		log
	end
end

# Include this module into a class that defines #each_child. It should
# maybe use #each instead, but its easier to be more specific, and use
# an alias.
#
# I don't want to force the class to cache children (eg where children
# are loaded on request in pst), because that forces the whole tree to
# be loaded. So, the methods should only call #each_child once, and 
# breadth first iteration holds its own copy of the children around.
#
# Main methods are #recursive, and #to_tree
module RecursivelyEnumerable
	def each_recursive_depth_first(&block)
		each_child do |child|
			yield child
			if child.respond_to? :each_recursive_depth_first
				child.each_recursive_depth_first(&block)
			end
		end
	end

	def each_recursive_breadth_first(&block)
		children = []
		each_child do |child|
			children << child if child.respond_to? :each_recursive_breadth_first
			yield child
		end
		children.each { |child| child.each_recursive_breadth_first(&block) }
	end

	def each_recursive mode=:depth_first, &block
		# we always actually yield ourself (the tree root) before recursing
		yield self
		send "each_recursive_#{mode}", &block
	end

	# the idea of this function, is to allow use of regular Enumerable methods
	# in a recursive fashion. eg:
	#
	#   # just looks at top level children
	#   root.find { |child| child.some_condition? }
	#   # recurse into all children getting non-folders, breadth first
	#   root.recursive(:breadth_first).select { |child| !child.folder? }
	#   # just get everything
	#   items = root.recursive.to_a
	#
	def recursive mode=:depth_first
		to_enum(:each_recursive, mode)
	end

	# streams a "tree" form of the recursively enumerable structure to +io+, or
	# return a string form instead if +io+ is not specified.
	#
	# mostly a debugging aid. can specify a different +method+ to call if desired,
	# though it should return a single line string.
	def to_tree io=nil, method=:inspect
		unless io
			to_tree io = StringIO.new
			return io.string
		end
		io << "- #{send method}\n"
		to_tree_helper io, method, '  '
	end

	def to_tree_helper io, method, prefix
		# i need to know when i get to the last child. use delay to know
		child = nil
		each_child do |next_child|
			if child
				io << "#{prefix}|- #{child.send method}\n"
				if child.respond_to? :to_tree_helper
					child.to_tree_helper io, method, prefix + '|  '
				end
			end
			child = next_child
		end
		# child is the last child
		io << "#{prefix}\\- #{child.send method}\n"
		if child.respond_to? :to_tree_helper
			child.to_tree_helper io, method, prefix + '   '
		end
	end
	protected :to_tree_helper
end

