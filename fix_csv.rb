#!/usr/bin/env ruby

text = File.read(ARGV[0])
fixed = text.gsub(/(?<!^)(?<!",)(?<!\d,)"(?!,")(?!,\d)(?!$)(?!,-\d)/, '""')
File.open(ARGV[0], "w") { |file| file.puts fixed }