require 'spreadsheet'
require 'fileutils'

unless File.exists?('recheck')		
	Dir.mkdir('recheck')
end

FileUtils.rm(Dir.glob("recheck/*.xls"))	

totalVariance = 0
totalDollarVariance = 0.0

puts "================================="
puts "==    BEGIN VARIENCE CHECK     =="
puts "================================="

Dir.glob("data_entry/*.xls").each do |file|

	puts "Processing file #{file}"

	book = Spreadsheet.open(file)
	
	recheck = []
	
	book.worksheet(0).each(1) do |row|
	
		variance = (row[3].to_i - row[4].to_i)
		dollarVariance = ((row[3].to_f * row[5].to_f) - (row[4].to_f * row[5].to_f))		
				
		if !variance.eql?(0) then
			recheck << {
				:binLocation => row[0],
				:itemNumber => row[1],
				:counted => row[3],
				:variance => variance,
				:dollarVariance => dollarVariance
			}
			
			totalVariance = totalVariance + variance
			totalDollarVariance = totalDollarVariance + dollarVariance
		end
	
	end
	
	if recheck.size > 0 then
		
		output = Spreadsheet::Workbook.new
				
		sheet1 = output.create_worksheet
		
		sheet1.row(0).height = 20
		
		sheet1[0,0] = "Bin Loc"
		sheet1[0,1] = "Item #"
		sheet1[0,2] = "Counted"
		sheet1[0,3] = "Variance"
		sheet1[0,4] = "Dollar Var."
		
		i = 1
		
		standard = Spreadsheet::Format.new()
		highlight = Spreadsheet::Format.new(:color => :red, :weight => :bold, :size => 14)
		
		sheet1.column(0).width = 15
		
		recheck.each do |item|
				
			sheet1.row(i).height = 20
			
			if item[:varience].to_i.abs > 10 or item[:dollarVariance].to_f.abs > 50.0
				sheet1.row(i).default_format = highlight
			else
				sheet1.row(i).default_format = standard
			end
		
			sheet1[i,0] = item[:binLocation]
			sheet1[i,1] = item[:itemNumber]
			sheet1[i,2] = item[:counted]
			sheet1[i,3] = item[:variance]
			sheet1[i,4] = item[:dollarVariance]			
			
			if item[:varience].to_i.abs > 10 or item[:dollarVariance].to_f.abs > 50.0
				sheet1[i,5] = "       RECHECK ______"
			end
		
			i = i.next
		
		end
				
		output.write("recheck/" << File.basename(file))
	
		puts "#{file}\t#{recheck.size}"
	
	end

end

puts "================================="
puts "==          FINISHED           =="
puts "================================="

puts "Total Variance: #{totalVariance}"
puts "Total Dollar Variance: $%0.2f" % totalDollarVariance


puts "================================="
puts "Press Enter To Close"

gets

