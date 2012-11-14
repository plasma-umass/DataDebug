#!/usr/local/bin/ruby

## NOTE: input21 should never have been fuzzed: it's a formula.
##       It is therefore excluded from the error test below.

require "rubygems"
require "csv"
require "colorize"

def print_color(colorize, val)
  if colorize then print "\"#{val}\", ".blue else print "\"#{val}\", " end
end

if ARGV[0] == nil
  puts "Usage:\truby fanniemae_diff.rb [colorized?] [just stats?] [max relative error?]"
  puts "\tE.g, ruby fanniemae_diff.rb true false"
  exit(1)
elsif ARGV[0].downcase == "true"
  puts ARGV[0]
  COLORIZE = true
else
  COLORIZE = false
end
if ARGV[1].nil? || ARGV[1].downcase == "false"
  NODATA = false
else
  NODATA = true
end
if ARGV[2].nil? || ARGV[2].downcase == "false"
  RELERR = false
else
  RELERR = true
end

FUZZDATA = "2012-11-13-mturk_fanniemae_fuzz.csv"

count = 0
errors = []

CSV.foreach(FUZZDATA) do |row|
  # we don't care about the first row-- just column headers
  if count == 0
    count += 1
    next
  end
  
  #"Input.col","Input.12","Input.13","Input.14","Input.15","Input.16","Input.18","Input.19","Input.21","Input.22","Input.26",
  #"Answer.12","Answer.13","Answer.14","Answer.15","Answer.16","Answer.18","Answer.19","Answer.21","Answer.22","Answer.26","Answer.col","Approve","Reject"
  
  true_col     = row[27]
  true_input12 = row[28]
  true_input13 = row[29]
  true_input14 = row[30]
  true_input15 = row[31]
  true_input16 = row[32]
  true_input18 = row[33]
  true_input19 = row[34]
  true_input21 = row[35]
  true_input22 = row[36]
  true_input26 = row[37]
  
  fuzz_input12 = row[38]
  fuzz_input13 = row[39]
  fuzz_input14 = row[40]
  fuzz_input15 = row[41]
  fuzz_input16 = row[42]
  fuzz_input18 = row[43]
  fuzz_input19 = row[44]
  fuzz_input21 = row[45]
  fuzz_input22 = row[46]
  fuzz_input26 = row[47]
  
  fails_input12 = false
  fails_input13 = false
  fails_input14 = false
  fails_input15 = false
  fails_input16 = false
  fails_input18 = false
  fails_input19 = false
  fails_input21 = false
  fails_input22 = false
  fails_input26 = false
  
  fails_input12 = true if true_input12 != fuzz_input12
  fails_input13 = true if true_input13 != fuzz_input13
  fails_input14 = true if true_input14 != fuzz_input14
  fails_input15 = true if true_input15 != fuzz_input15
  fails_input16 = true if true_input16 != fuzz_input16
  fails_input18 = true if true_input18 != fuzz_input18
  fails_input19 = true if true_input19 != fuzz_input19
  fails_input21 = true if true_input21 != fuzz_input21
  fails_input22 = true if true_input22 != fuzz_input22
  fails_input26 = true if true_input26 != fuzz_input26
  
  if fails_input12 || fails_input13 || fails_input14 || fails_input15 || fails_input16 || fails_input19 || fails_input22 || fails_input26
    error_r = { :true_col => true_col,
                :true_input12 => true_input12,
                :true_input13 => true_input13,
                :true_input14 => true_input14,
                :true_input15 => true_input15,
                :true_input16 => true_input16,
                :true_input18 => true_input18,
                :true_input19 => true_input19,
                :true_input21 => true_input21,
                :true_input22 => true_input22,
                :true_input26 => true_input26,
                :fuzz_input12 => fuzz_input12,
                :fuzz_input13 => fuzz_input13,
                :fuzz_input14 => fuzz_input14,
                :fuzz_input15 => fuzz_input15,
                :fuzz_input16 => fuzz_input16,
                :fuzz_input18 => fuzz_input18,
                :fuzz_input19 => fuzz_input19,
                :fuzz_input21 => fuzz_input21,
                :fuzz_input22 => fuzz_input22,
                :fuzz_input26 => fuzz_input26,
                :fails_input12 => fails_input12,
                :fails_input13 => fails_input13,
                :fails_input14 => fails_input14,
                :fails_input15 => fails_input15,
                :fails_input16 => fails_input16,
                :fails_input18 => fails_input18,
                :fails_input19 => fails_input19,
                :fails_input21 => fails_input21,
                :fails_input22 => fails_input22,
                :fails_input26 => fails_input26 }
    errors << error_r
  end

  count += 1
end

puts "Total Rows: " + count.to_s
puts "Total Errors: " + errors.size.to_s
puts "Error Rate: " + (errors.size.to_f / count.to_f * 100).to_s + "%"

count = 0
unless NODATA
  errors.each do |e|
    next if e[:fuzz_input12].empty? || e[:fuzz_input13].empty? || e[:fuzz_input14].empty? || e[:fuzz_input15].empty? || e[:fuzz_input16].empty? || e[:fuzz_input19].empty? || e[:fuzz_input22].empty? || e[:fuzz_input26].empty?
    print "#{count}: "
    if COLORIZE
      print "\"#{e[:true_col]}\","
      print_color(e[:fails_input12] && COLORIZE, e[:true_input12])
      print_color(e[:fails_input13] && COLORIZE, e[:true_input13])
      print_color(e[:fails_input14] && COLORIZE, e[:true_input14])
      print_color(e[:fails_input15] && COLORIZE, e[:true_input15])
      print_color(e[:fails_input16] && COLORIZE, e[:true_input16])
      print_color(e[:fails_input18] && COLORIZE, e[:true_input18])
      print_color(e[:fails_input19] && COLORIZE, e[:true_input19])
      print_color(e[:fails_input21] && COLORIZE, e[:true_input21])
      print_color(e[:fails_input22] && COLORIZE, e[:true_input22])
      print_color(e[:fails_input26] && COLORIZE, e[:true_input26])
      print " => ".red
    end
    print "\"#{e[:true_col]}\","
    print_color(e[:fails_input12] && COLORIZE, e[:fuzz_input12])
    print_color(e[:fails_input13] && COLORIZE, e[:fuzz_input13])
    print_color(e[:fails_input14] && COLORIZE, e[:fuzz_input14])
    print_color(e[:fails_input15] && COLORIZE, e[:fuzz_input15])
    print_color(e[:fails_input16] && COLORIZE, e[:fuzz_input16])
    print_color(e[:fails_input18] && COLORIZE, e[:fuzz_input18])
    print_color(e[:fails_input19] && COLORIZE, e[:fuzz_input19])
    print_color(e[:fails_input21] && COLORIZE, e[:fuzz_input21])
    print_color(e[:fails_input22] && COLORIZE, e[:fuzz_input22])
    print_color(e[:fails_input26] && COLORIZE, e[:fuzz_input26])
    print "\n"
    count += 1
  end
end

def mag_err(tval, fval)
  ((fval.to_f - tval.to_f) / tval.to_f).abs
end

count = 0
if RELERR
  errors.each do |e|
    next if e[:fuzz_input12].empty? || e[:fuzz_input13].empty? || e[:fuzz_input14].empty? || e[:fuzz_input15].empty? || e[:fuzz_input16].empty? || e[:fuzz_input19].empty? || e[:fuzz_input22].empty? || e[:fuzz_input26].empty?
    max = 0
    if e[:fails_input12]
      mag = mag_err(e[:true_input12].to_f,e[:fuzz_input12].to_f)
      if mag > max
        max = mag
      end
    end
    if e[:fails_input13]
      mag = mag_err(e[:true_input13].to_f,e[:fuzz_input13].to_f)
      if mag > max
        max = mag
      end
    end
    if e[:fails_input14]
      mag = mag_err(e[:true_input14].to_f,e[:fuzz_input14].to_f)
      if mag > max
        max = mag
      end
    end
    if e[:fails_input15]
      mag = mag_err(e[:true_input15].to_f,e[:fuzz_input15].to_f)
      if mag > max
        max = mag
      end
    end
    if e[:fails_input16]
      mag = mag_err(e[:true_input16].to_f,e[:fuzz_input16].to_f)
      if mag > max
        max = mag
      end
    end
    if e[:fails_input19]
      mag = mag_err(e[:true_input19].to_f,e[:fuzz_input19].to_f)
      if mag > max
        max = mag
      end
    end
    if e[:fails_input22]
      mag = mag_err(e[:true_input22].to_f,e[:fuzz_input22].to_f)
      if mag > max
        max = mag
      end
    end
    if e[:fails_input26]
      mag = mag_err(e[:true_input26].to_f,e[:fuzz_input26].to_f)
      if mag > max
        max = mag
      end
    end
    puts count.to_s + ": maximum % change in absolute error = " + mag.to_s
    count += 1
  end
end