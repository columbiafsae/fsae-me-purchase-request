require 'axlsx'
require 'bigdecimal'

def generate_form(data, path)

p = Axlsx::Package.new
wb = p.workbook

highlight = nil
wb.styles do |s|
  highlight = s.add_style :bg_color => "ffff00"
end

wb.add_worksheet(:name => "Purchase Request") do |ws|
  ws.add_row ["VENDOR"], :style => highlight
  ws.add_row ["Company:", data.vendor.name, "", "Requested By:", data.requested_by], :style => [highlight, nil, nil, highlight, highlight]
  ws.add_row ["Address:", data.vendor.address.join(", ")], :style => [highlight, nil]
  ws.add_row
  ws.add_row ["Phone:", data.vendor.phone, "", "Account No:", data.account], :style => [highlight, nil, nil, highlight]
  ws.add_row ["Fax:", data.vendor.fax], :style => [highlight, nil]
  ws.add_row ["Website:", data.vendor.url], :style => [highlight, nil]
  
  ws.add_row
  ws.add_row
  
  ws.add_row ["Item", "Part #", "Description", "Unit", "Chemical?", "Hazardous?", "Qty", "Unit Price", "Extended Price"]
  row_num = 1
  last_cell = nil
  for item in data.items do
    row_data = [row_num, item.part_number, item.desc, item.unit, item.chemical? ? "X" : "", item.hazardous? ? "X" : "", item.qty, item.price_unit]
    row = ws.add_row row_data
    ext_row = row.index + 1
    ext_col_A = 'A'.ord + row_data.length - 1
    ext_col_B = ext_col_A - 1
    last_cell = row.add_cell "=#{ext_col_A.chr}#{ext_row} * #{ext_col_B.chr}#{ext_row}"
    
    row_num += 1
  end
  
end

wb.use_shared_strings = true
p.serialize path

end

OrderItem = Struct.new :part_number, :desc, :url, :unit, :qty, :price_unit, :chemical?, :hazardous?
VendorData = Struct.new :name, :address, :phone, :fax, :url
OrderData = Struct.new :vendor, :date, :requested_by, :account, :items, :notes, :ship_to

item1 = OrderItem.new "1234", "Widget", "http://google.com", "each", "25", "132.11", false, false
vendor = VendorData.new "Bob's Widget Factory", ["2920 Broadway", "New York, NY 10027"], "510 545 3860", "510 545 3860", "http://kevinchen.co/"
order = OrderData.new vendor, "05-10-2014", "Kevin", "abcdefg", [item1], "", ["Vinod Nimmagadda", "Formula SAE", "500 W 120th St", "New York NY 10027"]

generate_form order, "out.xlsx"
