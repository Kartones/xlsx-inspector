# encoding: utf-8

require 'zip'

module XLSXInspector
	class Inspector

		public

		TEMP_FOLDER_PREFIX = 'xlsxi-'

		def initialize()
	  end

	  def inspect(xlsx_filename)
  		xml_block = ''
	  	dimension = nil
	  	temp_folder = get_temporally_folder()

    	unpacked_file_path = unpack(xlsx_filename, temp_folder)
    	if (!unpacked_file_path.nil?)
  	  	File.open(unpacked_file_path, "rb") do |file|
	      	xml_block = file.read(512)
	    	end
	    	dimension = find_dimension(xml_block)
	    end

	  	FileUtils.remove_entry_secure(temp_folder)

	  	dimension.nil? ? nil : calculate_size(dimension)
	  end

		private

	  def get_temporally_folder()
	    Dir.mktmpdir(TEMP_FOLDER_PREFIX)
	  end

	  def unpack(filename, temp_folder, zip_path=/xl\/worksheets\/sheet1.xml/)
	  	unpacked_file_path = nil
			Zip::File.open(filename) do |zip_file|
				zip_file.entries.each do |file_path|
					if !!(zip_path =~ file_path.name)
						unpacked_file_path = File.join(temp_folder, File.basename(file_path.name))
						File.open(unpacked_file_path,"wb") do |file|
							file.write(zip_file.read(file_path.name))
						end
					end
				end
		  end
		  unpacked_file_path
	  end

	  def find_dimension(xml_block)
	  	matching = /<dimension ref="([0-9A-Z:]+)"\/>/.match(xml_block)
	  	matching.size() == 2 ? matching[1] : nil
	  end

	  def calculate_size(dimension)
	  	alphabet_conversion = {}
	  	index = 0
	  	('A'..'Z').each do |x|
	  		index += 1
	  		alphabet_conversion[x.to_sym] = index
	  	end

	  	topleft, bottomright = dimension.split(':')
	  	first_column = /([A-Z]+)/.match(topleft)[1]
	  	first_row = /([0-9]+)/.match(topleft)[1].to_i

	  	last_column = /([A-Z]+)/.match(bottomright)[1]
	  	last_row = /([0-9]+)/.match(bottomright)[1].to_i

			columns = 1 + get_numerical_value(last_column, alphabet_conversion) - get_numerical_value(first_column, alphabet_conversion)
			rows = 1 + last_row - first_row

			columns * rows
	  end

	  def get_numerical_value(letters, alphabet_conversion)
	  	accumulator = -1
	  	position = 0
	  	while (letters.size() > 0)
	  		accumulator += alphabet_conversion[letters[-1].to_sym] + (26 ** position)
	  		letters.chop!
	  		position += 1
			end
			accumulator
	  end

	end
end