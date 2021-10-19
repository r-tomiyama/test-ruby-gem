require 'bundler/inline'

gemfile do
  source 'https://rubygems.org'
  gem 'rubyXL', require: false
end

require 'rubyxl'

# https://github.com/weshatheleopard/rubyXL/wiki/How-to#create-a-sheet-with-locked-and-unlocked-cells

class RubyXL::Cell
  def unlock
    xf = get_cell_xf.dup
    xf.protection = xf.protection&.dup || RubyXL::Protection.new
    xf.protection.locked = false
    xf.apply_protection = true
    self.style_index = workbook.register_new_xf(xf)
  end
end

@workbook = RubyXL::Workbook.new
@sheet = @workbook.first
@sheet.add_cell(0, 0, "Hello")
@sheet.add_cell(0, 1, "World")
@sheet.add_cell(1, 0, "Hello")
@sheet.add_cell(1, 1, "World")
@sheet[1][1].unlock
@sheet.sheet_protection = RubyXL::WorksheetProtection.new(sheet: true)
@workbook.write("out.xlsx")
`open out.xlsx`
