#!/usr/bin/env ruby

require 'roo'

def filter_rows(filename, sheet_index)
  xlsx = Roo::Excelx.new(filename)
  file = File.open("./accounts.csv", "w")

  xlsx.sheet(sheet_index).select do |row|
    row[2].length == 10 || row[2].length == 11
  end.map do |row|
    row[2] = row[2][1..-1] if row[2].length > 10
    file.write("#{row[0]},#{row[2]},#{row[3]}\n")
  end
  file.close
end

filter_rows "./ekTally.xls.xlsx", 0
