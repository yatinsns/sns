#!/usr/bin/env ruby

require 'roo'

COLORS = ["TURKISH", "TELIA", "MOUSE", "CREAM", "GOLDEN", "GRAY", "BLACK", "WHITE", "ORANGE", "BLUE", "COFFEE", "MEHROON", "YELLOW", "BOTTLE GREEN", "RED", "PINK", "FIROZI", "PURPLE", "BROWN", "MIX"]

def filter_rows(filename, sheet_index)
  xlsx = Roo::Excelx.new(filename)
  file = File.open("./products.csv", "w")

  s_no = 1
  xlsx.sheet(sheet_index).each_with_index do |row, index|
    if index == 0
      file.write("S.NO,PRODUCT NAME,FOR,FABRIC,SIZE,COLORS,WSP,MRP,PRODUCT_TYPE,FILTERS\n")
    elsif row[1] == "AFZAL"
      name = row[3]
      type = "Men"
      if name =~ /CHILD/i || name =~ /GIRL/i
	type = "Kids"
      elsif name =~ /LADIES/i
        type = "Women"
      end

      size = row[4].to_s
      size = "Free" if size == "F"
      size = "Free" if size.length == 0
      size = size.gsub("3XL", "XXXL")
      size = size.gsub("4XL", "XXXXL")
      size = size.gsub("''", "")

      fabric = ""

      if name =~ /T\/L/
	fabric = "Treslon"
      elsif name =~ /NS\b/
        fabric = "Nylon Seasor"
      elsif name =~ /H\/PU/
	fabric = "Honda PU"
      elsif name =~ /Cotton/
	fabric = "Cotton"
      end

      filters = Array.new
      if name =~ /S\/L/i
	filters << "Sleeveless"
      end
      if name =~ /R\/S/i
	filters << "Reversible"
      end

      product = "Jackets"
      if name =~ /W\/C/i
	product = "Windcheaters"
      end

      colors = []
      (5...25).each do |column_index|
        color = COLORS[column_index - 5]
        count = row[column_index]
	colors << color unless count.nil?
      end


      file.write("#{s_no},#{row[3]},#{type},#{fabric},#{size},#{colors.join(":")},#{row[26]},#{row[27]},#{product},#{filters.join(":")}\n")
      s_no = s_no + 1 
    end
  end
  file.close
end

filter_rows "./SNS.xlsx", 0
