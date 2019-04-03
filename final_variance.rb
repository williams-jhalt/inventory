require 'spreadsheet'
require 'fileutils'

totalVariance = 0
totalDollarVariance = 0.0

puts "================================="
puts "==    BEGIN VARIENCE CHECK     =="
puts "================================="
	
recheck = []

Dir.glob("done/*.xls").each do |file|

	puts "Processing file #{file}"

	book = Spreadsheet.open(file)
	
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

end

puts "================================="
puts "==    BEGIN VARIENCE CHECK     =="
puts "================================="
	
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
		
	sheet1.column(0).width = 15
	
	recheck.each do |item|
			
		sheet1.row(i).height = 20
	
		sheet1[i,0] = item[:binLocation]
		sheet1[i,1] = item[:itemNumber]
		sheet1[i,2] = item[:counted]
		sheet1[i,3] = item[:variance]
		sheet1[i,4] = item[:dollarVariance]
	
		i = i.next
	
	end
			
	output.write("final_variance.xls")

	puts "#{file}\t#{recheck.size}"

end

puts "================================="
puts "==          FINISHED           =="
puts "================================="

puts "Total Variance: #{totalVariance}"
puts "Total Dollar Variance: $%0.2f" % totalDollarVariance


puts "================================="
puts "Press Enter To Close"

gets

