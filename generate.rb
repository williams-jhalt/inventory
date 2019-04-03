require 'writeexcel'
require 'csv'

unless File.exists?('worksheets')
	Dir.mkdir('worksheets')
end

unless File.exists?('data_entry')
	Dir.mkdir('data_entry')
end

puts "================================="
puts "==       LOADING ITEMS         =="
puts "================================="

bins = Hash.new

CSV.foreach("items.csv") do |row| 

	itemNumber = row[0]
	name = row[1]
	unitOfMeasure = row[2]
	binLocation = row[3]
	onHand = row[4]
	price = row[5]
	
	/(?<bin>\w+)(-.*)+/ =~ binLocation
	
	unless bins.key?(bin) then bins[bin] = Array.new end
	
	bins[bin] << {
		:itemNumber => itemNumber,
		:name => name,
		:unitOfMeasure => unitOfMeasure,
		:binLocation => binLocation,
		:onHand => onHand,
		:price => price
	}

end

puts "Finished: found #{bins.length} locations"

puts "================================="
puts "==     CREATING WORKSHEETS     =="
puts "================================="

# create worksheets
bins.each do |bin, items|

	items.sort_by! { |x| x[:binLocation] }
	
	if bin.nil? or bin.empty?
		filename = "worksheets/unknown.xls"
	else 
		filename = "worksheets/#{bin}.xls"
	end
	
	puts filename

	workbook = WriteExcel.new(filename)
	
	title_format = workbook.add_format(:center_across => 1, :bold => 1, :size => 14)	
	header_format = workbook.add_format(:size => 10, :bold => 1, :align => 'center')	
	cell_format = workbook.add_format(:size => 11, :border => 1, :font => 'Calibri')
	
	worksheet = workbook.add_worksheet
	worksheet.hide_gridlines
	worksheet.set_landscape
	worksheet.repeat_rows(0,3)
	worksheet.set_footer("  OS_________    REG_________   Page &P of &N    DE_________")
	
	worksheet.set_column(0, 0, 10)
	worksheet.set_column(1, 0, 13)
	worksheet.set_column(2, 0, 20)
	worksheet.set_column(3, 0, 10)
	worksheet.set_column(4, 0, 13)
	worksheet.set_column(5, 0, 35)
	worksheet.set_column(6, 0, 7)
	worksheet.set_column(7, 0, 7)
	
	worksheet.set_row(0, 20)
	worksheet.write(0, 0, "WILLIAMS TRADING CO - INVENTORY", title_format)
	worksheet.write_blank(0, 1, title_format)
	worksheet.write_blank(0, 2, title_format)
	worksheet.write_blank(0, 3, title_format)
	worksheet.write_blank(0, 4, title_format)
	worksheet.write_blank(0, 5, title_format)
	worksheet.write_blank(0, 6, title_format)
	worksheet.write_blank(0, 7, title_format)
		
	worksheet.set_row(1, 20)
	worksheet.write(1, 0, bin, title_format)
	worksheet.write_blank(1, 1, title_format)
	worksheet.write_blank(1, 2, title_format)
	worksheet.write_blank(1, 3, title_format)
	worksheet.write_blank(1, 4, title_format)
	worksheet.write_blank(1, 5, title_format)
	worksheet.write_blank(1, 6, title_format)
	worksheet.write_blank(1, 7, title_format)	
	
	worksheet.set_row(2, 20)
	worksheet.write_blank(2, 0)
	
	worksheet.set_row(3, 20)
	worksheet.write(3, 0, "BIN", header_format)
	worksheet.write(3, 1, "ITEM #", header_format)
	worksheet.write(3, 2, "OVERSTOCK", header_format)
	worksheet.write(3, 3, "REG STK", header_format)
	worksheet.write(3, 4, "BIN", header_format)
	worksheet.write(3, 5, "DESCRIPTION", header_format)
	worksheet.write(3, 6, "U OF M", header_format)
	worksheet.write(3, 7, "TOT", header_format)
	
	worksheet.write()
	
	i = 4
	
	items.each do |item|
	
		worksheet.set_row(i, 20)
		worksheet.write(i, 0, item[:binLocation], cell_format)
		worksheet.write(i, 1, item[:itemNumber], cell_format)
		worksheet.write_blank(i, 2, cell_format)
		worksheet.write_blank(i, 3, cell_format)
		worksheet.write(i, 4, item[:binLocation], cell_format)
		worksheet.write(i, 5, item[:name], cell_format)
		worksheet.write(i, 6, item[:unitOfMeasure], cell_format)
		worksheet.write_blank(i, 7, cell_format)
		
		i = i.next
	
	end
	
	workbook.close

end

puts "Finished!"

puts "================================="
puts "== CREATING DATA ENTRY SHEETS  =="
puts "================================="

# create data entry sheets
bins.each do |bin, items|

	items.sort_by! { |x| x[:binLocation] }
	
	if bin.nil? or bin.empty?
		filename = "data_entry/unknown.xls"
	else 
		filename = "data_entry/#{bin}.xls"
	end
	
	puts filename

	workbook = WriteExcel.new(filename)
	
	header_format = workbook.add_format(:size => 10, :bold => 1, :align => 'center', :locked => 1)	
	cell_format = workbook.add_format(:size => 11, :border => 1, :font => 'Calibri', :locked => 1)
	edit_cell_format = workbook.add_format(:size => 11, :border => 1, :font => 'Calibri', :locked => 0)
	currency_format = workbook.add_format(:size => 11, :border => 1, :font => 'Calibri', :num_format => '$0.00', :locked => 1)
	
	worksheet = workbook.add_worksheet
	worksheet.hide_gridlines
	worksheet.set_landscape
	worksheet.protect
	
	worksheet.set_column(0, 0, 10)
	worksheet.set_column(1, 0, 13)
	worksheet.set_column(2, 0, 40)
	worksheet.set_column(3, 0, 10)
	worksheet.set_column(4, 0, 10)
	worksheet.set_column(5, 0, 10)
	worksheet.set_column(6, 0, 13)
	worksheet.set_column(7, 0, 13)
		
	worksheet.set_row(0, 20)
	worksheet.write(0, 0, "BIN", header_format)
	worksheet.write(0, 1, "ITEM #", header_format)
	worksheet.write(0, 2, "NAME", header_format)
	worksheet.write(0, 3, "COUNT", header_format)
	worksheet.write(0, 4, "ON HAND", header_format)
	worksheet.write(0, 5, "COST", header_format)
	worksheet.write(0, 6, "VAR. COUNT", header_format)
	worksheet.write(0, 7, "VAR. DOLLAR", header_format)
	
	worksheet.write()
	
	i = 1
	
	items.each do |item|
	
		xlRow = i + 1
	
		worksheet.set_row(i, 20)
		worksheet.write(i, 0, item[:binLocation], cell_format)
		worksheet.write(i, 1, item[:itemNumber], cell_format)
		worksheet.write(i, 2, item[:name], cell_format)
		worksheet.write(i, 3, item[:onHand].to_i, edit_cell_format)
		worksheet.write(i, 4, item[:onHand].to_i, cell_format)
		worksheet.write(i, 5, item[:price].to_f, currency_format)
		worksheet.write_formula(i, 6, "=D#{xlRow}-E#{xlRow}", cell_format)
		worksheet.write_formula(i, 7, "=(D#{xlRow}*F#{xlRow})-(E#{xlRow}*F#{xlRow})", currency_format)
		
		i = i.next
	
	end
	
	worksheet.set_row(i, 20)
	worksheet.write_formula(i, 6, "=SUM(G2:G#{i})", cell_format)
	worksheet.write_formula(i, 7, "=SUM(H2:H#{i})", currency_format)
	
	workbook.close

end

puts "Finished!"

puts "================================="
puts "==    CREATING MASTER LIST     =="
puts "================================="

# create master bin list
workbook = WriteExcel.new("worksheets/master_bin_list.xls")

title_format = workbook.add_format(:center_across => 1, :bold => 1, :size => 14)	
header_format = workbook.add_format(:size => 10, :bold => 1, :align => 'center')	
cell_format = workbook.add_format(:size => 11, :border => 1, :font => 'Calibri')
header_rot_format = workbook.add_format(:size => 10, :bold => 1, :rotation => 90, :align => 'center')

worksheet = workbook.add_worksheet
worksheet.hide_gridlines
worksheet.set_landscape
	
worksheet.set_column(0, 0, 30)
worksheet.set_column(1, 0, 30)
worksheet.set_column(2, 0, 13)
worksheet.set_column(3, 0, 13)
worksheet.set_column(4, 0, 13)
worksheet.set_column(5, 0, 4)
worksheet.set_column(6, 0, 4)
worksheet.set_column(7, 0, 4)
worksheet.set_column(8, 0, 4)

worksheet.repeat_rows(0,2)
worksheet.set_footer("Page &P of &N")
	
worksheet.set_row(0, 20)
worksheet.write(0, 0, "WILLIAMS TRADING CO - INVENTORY", title_format)
worksheet.write_blank(0, 1, title_format)
worksheet.write_blank(0, 2, title_format)
worksheet.write_blank(0, 3, title_format)
worksheet.write_blank(0, 4, title_format)
worksheet.write_blank(0, 5, title_format)
worksheet.write_blank(0, 6, title_format)
worksheet.write_blank(0, 7, title_format)
worksheet.write_blank(0, 8, title_format)

worksheet.set_row(1, 20)
worksheet.write(1, 0, "MASTER BIN LIST", title_format)
worksheet.write_blank(1, 1, title_format)
worksheet.write_blank(1, 2, title_format)
worksheet.write_blank(1, 3, title_format)
worksheet.write_blank(1, 4, title_format)
worksheet.write_blank(1, 5, title_format)
worksheet.write_blank(1, 6, title_format)
worksheet.write_blank(1, 7, title_format)
worksheet.write_blank(1, 8, title_format)
	
worksheet.set_row(2, 100)
worksheet.write(2, 0, "O.S. Sign-in", header_format)
worksheet.write(2, 1, "Reg. Stock Sign-in", header_format)
worksheet.write(2, 2, "Start Bin", header_format)
worksheet.write(2, 3, "End Bin", header_format)
worksheet.write(2, 4, "File Name", header_format)
worksheet.write(2, 5, " Count Wks. Print", header_rot_format)
worksheet.write(2, 6, " Input Count", header_rot_format)
worksheet.write(2, 7, " Variance Print", header_rot_format)
worksheet.write(2, 8, " Import File", header_rot_format)

i = 3

bins.keys.compact.sort.each do |bin|

	items = bins[bin]
		
	items.sort_by! { |x| x[:binLocation] }
	
	worksheet.set_row(i, 20)
	worksheet.write_blank(i, 0, cell_format)
	worksheet.write_blank(i, 1, cell_format)	
	worksheet.write(i, 2, items.first[:binLocation], cell_format)
	worksheet.write(i, 3, items.last[:binLocation], cell_format)
	worksheet.write(i, 4, bin, cell_format)
	worksheet.write_blank(i, 5, cell_format)
	worksheet.write_blank(i, 6, cell_format)
	worksheet.write_blank(i, 7, cell_format)
	worksheet.write_blank(i, 8, cell_format)
	
	i = i.next
	
end
	
workbook.close()	
worksheet.write()

puts "Finished!"

puts "================================="
puts "Press Enter To Close"

gets
