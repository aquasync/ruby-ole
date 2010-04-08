warn <<-end
Use of ole/file_system is deprecated. Use ole/storage/file_system
or better just ole/storage (file_system is recommended api and is
enabled by default).
end

require 'ole/storage/file_system'
