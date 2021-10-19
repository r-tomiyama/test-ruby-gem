require 'bundler/inline'

gemfile do
  source 'https://rubygems.org'
  gem 'rubyXL', require: false
end

require 'rubyxl'
require 'rubyXL/convenience_methods'

# docs
# https://github.com/weshatheleopard/rubyXL/issues/272
# http://www.datypic.com/sc/ooxml/e-ssml_outlinePr-1.html

@workbook = RubyXL::Workbook.new
@sheet = @workbook.first

(1..15).each_with_index do |num, row_index|
  ("A".."G").each_with_index do |v, column_index|
    @sheet.add_cell(row_index, column_index, "#{v}#{num}")
  end
end

@sheet[1].outline_level = 10
@sheet[2].outline_level = 9
@sheet[3].outline_level = 8
@sheet[4].outline_level = 7
@sheet[5].outline_level = 6
@sheet[6].outline_level = 5
@sheet[7].outline_level = 4
@sheet[8].outline_level = 3
@sheet[9].outline_level = 2
@sheet[10].outline_level = 1

@sheet.cols.get_range(1).outline_level = 2
@sheet.cols.get_range(2).outline_level = 1

@sheet.sheet_pr = RubyXL::WorksheetProperties.new()
@sheet.sheet_pr.outline_pr = RubyXL::OutlineProperties.new()

@workbook.write("out.xlsx")
`open out.xlsx`
