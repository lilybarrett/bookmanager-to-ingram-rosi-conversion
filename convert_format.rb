require 'roo'
require 'roo-xls'
require 'write_xlsx'
require 'byebug'

# Read the structure of Spreadsheet A
def analyze_structure(spreadsheet_a_path)
  spreadsheet_a = Roo::Spreadsheet.open(spreadsheet_a_path)
  sheet = spreadsheet_a.sheet(1)
  # Extract column headers (assume headers are in the first row)
  column_names_a = sheet.row(2)
  column_names_a
end

# Convert Spreadsheet B to match the structure of Spreadsheet A
def convert_format(spreadsheet_b_path, column_names_a, output_spreadsheet_path)
  spreadsheet_b = Roo::Spreadsheet.open(spreadsheet_b_path, extension: :xls)
  sheet_b = spreadsheet_b.sheet(0)
  
  # Extract column headers from Spreadsheet B
  column_names_b = sheet_b.row(1)
  
  # Create new spreadsheet C with the same format as Spreadsheet A
  workbook = WriteXLSX.new(output_spreadsheet_path)
  worksheet = workbook.add_worksheet
  
  # Write column names of Spreadsheet A to the new spreadsheet (first row)
  column_names_a.each_with_index do |col_name, index|
    worksheet.write(0, index, col_name)
  end

  column_names_mapping = {
    "ISBN" => "EAN",
    "Qty" => "Ord Qty",
    "Title" => "Title",
    "Author" => "Author",
    "PubDate" => "Pub Date",
    "Disc" => "Disc",
    "Retail" => "Retail Price"
  }

  column_mapping = {}
  column_names_mapping.each do |col_b_name, col_a_name|
    col_b_index = column_names_b.index(col_b_name)  # Find the index of the column in B
    col_a_index = column_names_a.index(col_a_name)  # Find the index of the column in A
    column_mapping[col_b_index] = col_a_index if col_b_index && col_a_index  # Store only if both columns exist
  end

  # Write data from Spreadsheet B to the new spreadsheet, in the same format as Spreadsheet A
  (2..sheet_b.last_row).each do |row_index|
    column_mapping.each do |col_b_index, col_a_index|
      if col_b_index && col_a_index
        # If it's a date, convert to M/D/Y format
        # byebug
        if column_names_a[col_a_index] == "Pub Date"
          date = Date.parse(sheet_b.cell(row_index, col_b_index + 1))
          worksheet.write(row_index - 1, col_a_index, date.strftime("%-m/%-d/%Y"))  # sheet_b.cell is 1-based
          next
        end
        # if discount is 40, input the string "REG"
        if column_names_a[col_a_index] == "Disc"
          discount = sheet_b.cell(row_index, col_b_index + 1)
          if discount == 40
            worksheet.write(row_index - 1, col_a_index, "REG")
          else
            worksheet.write(row_index - 1, col_a_index, discount)
          end
          next
        end
        # Write the cell value from Spreadsheet B to the correct column in Spreadsheet A
        worksheet.write(row_index - 1, col_a_index, sheet_b.cell(row_index, col_b_index + 1))  # sheet_b.cell is 1-based
      else
        # Leave blank if column mapping does not exist
        worksheet.write(row_index - 1, col_a_index, "")
      end
    end
  end
  
  # Write data from Spreadsheet B to the new spreadsheet, in the same format as Spreadsheet A
  # For any non-matching columns, leave the cell blank

  # Close and save the new spreadsheet
  workbook.close
end

# Paths to the Excel files
spreadsheet_a_path = 'spreadsheets/SpreadsheetA.xlsx'
spreadsheet_b_path = 'spreadsheets/SpreadsheetB.xls'
output_spreadsheet_path = 'spreadsheets/SpreadsheetC.xlsx'

# Analyze Spreadsheet A's structure
column_names_a = analyze_structure(spreadsheet_a_path)

# Convert Spreadsheet B to match Spreadsheet A's format
convert_format(spreadsheet_b_path, column_names_a, output_spreadsheet_path)

puts "Spreadsheet B has been converted to match Spreadsheet A's format in #{output_spreadsheet_path}"
