require 'bundler/inline'

gemfile do
  source 'https://rubygems.org'
  gem 'rubyXL', require: false
end

require 'rubyxl'

@workbook = RubyXL::Workbook.new
@sheet = @workbook.first
@sheet.add_cell(0, 0, "Hello World")
@workbook.write("out.xlsx")
`open out.xlsx`
