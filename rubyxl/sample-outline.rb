require 'bundler/inline'

gemfile do
  source 'https://rubygems.org'
  gem 'rubyXL', require: false
end

require 'rubyxl'

# docs
# https://github.com/weshatheleopard/rubyXL/issues/272
# http://www.datypic.com/sc/ooxml/e-ssml_outlinePr-1.html

@workbook = RubyXL::Workbook.new
@sheet = @workbook.first

(1..5).each_with_index do |num, row_index|
  ("A".."G").each_with_index do |v, column_index|
    @sheet.add_cell(row_index, column_index, "#{v}#{num}")
  end
end

@sheet[1].outline_level = 1
@sheet[2].outline_level = 2
@sheet[3].outline_level = 5 # 3になる

@sheet.cols.get_range(1).outline_level = 1
@sheet.cols.get_range(2).outline_level = 2
@sheet.cols.get_range(3).outline_level = 3

@sheet.cols.get_range(4).outline_level = 1
@sheet.cols.get_range(5).outline_level = 2
@sheet.cols.get_range(6).outline_level = 2

@sheet.cols.get_range(8).outline_level = 2 # 1になる

@sheet.sheet_pr = RubyXL::WorksheetProperties.new
@sheet.sheet_pr.outline_pr = RubyXL::OutlineProperties.new(
  summary_below: false, summary_right: false,
)
@sheet.sheet_format_pr = RubyXL::WorksheetFormatProperties.new(
  outline_level_row: 0, outline_level_col: 0
  # 効かない
)

@workbook.write("out.xlsx")
`open out.xlsx`
