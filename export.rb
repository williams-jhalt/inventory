require 'spreadsheet'
require 'fileutils'
require 'csv'

unless File.exists?('export')		
	Dir.mkdir('export')
end

FileUtils.rm(Dir.glob("export/*.csv"))	

puts "================================="
puts "==    BEGIN EXPORT FOR ERP     =="
puts "================================="

Dir.glob("done/*.xls").each do |file|

	puts "Processing file #{file}"

	book = Spreadsheet.open(file)
	
	CSV.open("export/" << File.basename(file, ".xls") << ".csv", "wb") do |csv|
		
		book.worksheet(0).each(1) do |row|	
		
			unless row[1].nil? then
		
				csv << [ row[1], row[3].to_i ]
				
			end
	
		end
	
	end

end

puts "================================="
puts "==          FINISHED           =="
puts "================================="
puts "Press Enter To Close"

gets

