require 'roo'
require 'roo-xls'
require 'write_xlsx'
require 'byebug'
require 'date'

# Created for Lovestruck Books & Cafe, a book-buying client preparing to open their store in late 2024.
# The client uses Bookmanager to select books for their store, including via Ingram.
# However, Ingram requires opening orders to be in their ROSI format in order to apply special discounts.
# This formula converts a spreadsheet of Ingram selections downloaded from Bookmanager to match the format of Ingram's ROSI spreadsheet,
# and cross-checks the ROSI for Ingram-platform-specific information like "Desire Status" and "Prod Type."

# Read the structure and data of Spreadsheet A, which is the ROSI format
def read_spreadsheet_a(spreadsheet_a_path)
  spreadsheet_a = Roo::Spreadsheet.open(spreadsheet_a_path)
  data = {}
  column_names_a = nil

  # Iterate through each sheet in the workbook
  spreadsheet_a.sheets.each do |sheet_name|
    next if sheet_name == "Summary"  # Skip the first sheet
    sheet = spreadsheet_a.sheet(sheet_name)

    # Extract column headers (assume headers are in the second row of each sheet)
    column_names_a ||= sheet.row(2)  # Assign only once from the first sheet

    ean_index = column_names_a.index("EAN")

    # Iterate through the rows of the current sheet and store additional data
    (3..sheet.last_row).each do |row_index|
      ean_value = sheet.cell(row_index, ean_index + 1)  # +1 because Roo is 1-based

      next if ean_value.nil?  # Skip if EAN/ISBN is missing

      data[ean_value] ||= {}  # Initialize hash for this EAN/ISBN

      column_names_a.each_with_index do |col_name, col_index|
        next if col_index == ean_index  # Skip the EAN column itself
        data[ean_value][col_name] = sheet.cell(row_index, col_index + 1)
      end
    end
  end

  { column_names_a: column_names_a, data: data }
end

# Convert Spreadsheet B, in the Bookmanager format, to match the structure of Spreadsheet A
def convert_format(spreadsheet_b_path, spreadsheet_a_data, output_spreadsheet_path)
  spreadsheet_b = Roo::Spreadsheet.open(spreadsheet_b_path, extension: :xls)
  sheet_b = spreadsheet_b.sheet(0)
  column_names_a = spreadsheet_a_data[:column_names_a]
  additional_item_data = spreadsheet_a_data[:data]
  
  # Create new spreadsheet C with the same format as Spreadsheet A
  workbook = WriteXLSX.new(output_spreadsheet_path)
  worksheet = workbook.add_worksheet
  
  # Write column names of Spreadsheet A to the new spreadsheet (first row)
  column_names_a.each_with_index do |col_name, index|
    worksheet.write(0, index, col_name)
  end

  # Mapping of Spreadsheet B columns to Spreadsheet A columns
  column_names_mapping = {
    "ISBN" => "EAN",
    "Qty" => "Ord Qty",
    "Title" => "Title",
    "Author" => "Author",
    "PubDate" => "Pub Date",
    "Disc" => "Disc",
    "Retail" => "Retail Price"
  }

  # Get the indices of the columns in Spreadsheet B for reference
  col_b_indices = {}
  sheet_b.row(1).each_with_index do |col_name, index|
    col_b_indices[col_name] = index
  end

  counter_found = 0
  counter_not_found = 0

  # Write data from Spreadsheet B to the new spreadsheet
  (2..sheet_b.last_row).each do |row_index|
    ean_value = sheet_b.cell(row_index, col_b_indices["ISBN"] + 1)
    
    # Write mapped columns
    column_names_mapping.each do |col_b_name, col_a_name|
      col_a_index = column_names_a.index(col_a_name)
      if col_a_index
        # Handle specific logic for Pub Date and Disc
        if col_a_name == "Pub Date"
          date = Date.parse(sheet_b.cell(row_index, col_b_indices[col_b_name] + 1))
          worksheet.write(row_index - 1, col_a_index, date.strftime("%-m/%-d/%Y"))
        elsif col_a_name == "Disc"
          discount = sheet_b.cell(row_index, col_b_indices[col_b_name] + 1)
          worksheet.write(row_index - 1, col_a_index, discount == 40 ? "REG" : discount)
        else
          worksheet.write(row_index - 1, col_a_index, sheet_b.cell(row_index, col_b_indices[col_b_name] + 1))
        end
      end
    end

    allowed_keys_to_write = ["Desire Status", "BISAC Description", "Section Description", "Prod Type", "Cost", "Total Cost"]

    # Append additional item data if EAN/ISBN exists
    if additional_item_data[ean_value]
      counter_found += 1
      additional_item_data[ean_value].each do |key, value|
        if allowed_keys_to_write.include?(key)
          col_index = column_names_a.index(key)  # Find the index in the A format
          worksheet.write(row_index - 1, col_index, value) if col_index  # Only write if column exists
        end
      end
    else
      counter_not_found += 1
      puts "Additional data not found for EAN/ISBN: #{ean_value}"
    end
  end

  puts "There are #{counter_found} items with additional data found in SpreadsheetA."
  puts "There are #{counter_not_found} items with additional data not found in SpreadsheetA."

  # Close and save the new spreadsheet
  workbook.close
end

# Paths to the Excel files
spreadsheet_a_path = 'spreadsheets/SpreadsheetA.xlsx'
spreadsheet_b_path = 'spreadsheets/SpreadsheetB.xls'
output_spreadsheet_path = 'spreadsheets/SpreadsheetC.xlsx'

# Read Spreadsheet A's structure and data
spreadsheet_a_data = read_spreadsheet_a(spreadsheet_a_path)

# Convert Spreadsheet B to match Spreadsheet A's format
convert_format(spreadsheet_b_path, spreadsheet_a_data, output_spreadsheet_path)

puts "Spreadsheet B has been converted to match Spreadsheet A's format in #{output_spreadsheet_path}"

# To run this script, you need to have the following gems installed:
# roo, roo-xls, write_xlsx
# You can install them using the following commands:
# gem install roo
# gem install roo-xls
# gem install write_xlsx

# Run the script using the following command:
# ruby convert_format.rb
