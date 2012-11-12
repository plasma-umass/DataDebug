#!/usr/local/bin/ruby

require "rubygems"
require "csv"
require "colorize"

if ARGV[0] == nil
  puts "Usage:\truby timecard_diff.rb [print numeric errors?]"
  puts "\tE.g, ruby timecard_diff.rb true"
  exit(1)
elsif ARGV[0].downcase == "true"
  puts ARGV[0]
  VERBOSE = true
else
  VERBOSE = false
end

FUZZDATA = "2012-11-11_timecard_human_fuzz.csv"

count = 0
errors = []

CSV.foreach(FUZZDATA) do |row|
  # we don't care about the first row-- just column headers
  if count == 0
    count += 1
    next
  end
  
  true_id   = row[27].to_i
  true_date = row[28]
  true_code = row[29].to_i
  true_name = row[30]
  true_unit = row[31].to_f
  fuzz_id   = row[32].to_i
  fuzz_date = row[33]
  fuzz_code = row[34].to_i
  fuzz_name = row[35]
  fuzz_unit = row[36].to_f
  
  fails_id   = false
  fails_date = false
  fails_code = false
  fails_name = false
  fails_unit = false
  
  fails_id = true if true_id != fuzz_id
  fails_date = true if true_date != fuzz_date
  fails_code = true if true_code != fuzz_code
  fails_name = true if true_name != fuzz_name
  fails_unit = true if true_unit != fuzz_unit
  
  if fails_id || fails_date || fails_code || fails_name || fails_unit
    error_r = { :fails_id => fails_id,
                :fails_date => fails_date,
                :fails_code => fails_code,
                :fails_name => fails_name,
                :fails_unit => fails_unit,
                :true_id => true_id,
                :true_date => true_date,
                :true_code => true_code,
                :true_name => true_name,
                :true_unit => true_unit,
                :fuzz_id => fuzz_id,
                :fuzz_date => fuzz_date,
                :fuzz_code => fuzz_code,
                :fuzz_name => fuzz_name,
                :fuzz_unit => fuzz_unit }
    errors << error_r
  end

  count += 1
end

puts "Total Rows: " + count.to_s
puts "Total Errors: " + errors.size.to_s
puts "Error Rate: " + (errors.size.to_f / count.to_f * 100).to_s + "%"

numeric_errors = errors.select { |e| e[:fails_id] || e[:fails_code] || e[:fails_unit] }

puts "Numeric Errors: " + numeric_errors.size.to_s
puts "Numeric Error Rate: " + (numeric_errors.size.to_f / count.to_f * 100).to_s + "%"

if VERBOSE
  numeric_errors.each do |e|
    if e[:fails_id] then print "\"#{e[:true_id]}\", ".blue else print "\"#{e[:true_id]}\", " end
    if e[:fails_date] then print "\"#{e[:true_date]}\", ".blue else print "\"#{e[:true_date]}\", " end
    if e[:fails_code] then print "\"#{e[:true_code]}\", ".blue else print "\"#{e[:true_code]}\", " end
    if e[:fails_name] then print "\"#{e[:true_name]}\", ".blue else print "\"#{e[:true_name]}\", " end
    if e[:fails_unit] then print "\"#{e[:true_unit]}\" ".blue else print "\"#{e[:true_unit]}\", " end
    print " => ".red
    if e[:fails_id] then print "\"#{e[:fuzz_id]}\", ".blue else print "\"#{e[:fuzz_id]}\", " end
    if e[:fails_date] then print "\"#{e[:fuzz_date]}\", ".blue else print "\"#{e[:fuzz_date]}\", " end
    if e[:fails_code] then print "\"#{e[:fuzz_code]}\", ".blue else print "\"#{e[:fuzz_code]}\", " end
    if e[:fails_name] then print "\"#{e[:fuzz_name]}\", ".blue else print "\"#{e[:fuzz_name]}\", " end
    if e[:fails_unit] then print "\"#{e[:fuzz_unit]}\", ".blue else print "\"#{e[:fuzz_unit]}\"" end
    print "\n"
  end
end

puts "BEGIN FUZZ OUTPUT:"
numeric_errors.each do |e|
  puts "\"#{e[:fuzz_id]}\", \"#{e[:fuzz_date]}\", \"#{e[:fuzz_code]}\", \"#{e[:fuzz_name]}\", \"#{e[:fuzz_unit]}\""
end