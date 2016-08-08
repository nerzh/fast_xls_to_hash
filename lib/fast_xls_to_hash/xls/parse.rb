require 'spreadsheet'

module FastXlsToHash

  class ParseXls

    FORMULA_ERROR = 'Error'

    def self.to_hash(file)
      book = Spreadsheet.open(file)
      hash = Hash.new{ |hash, key| hash[key] = {} }
      FastXlsToHash.with_xls_style ? with_characters(book, hash) : without_characters(book, hash)

      kill(book)

      hash
    end

    private

    def self.without_characters(book, hash)
      book.worksheets.each do |sheet|
        sheet.each_with_index do |row, row_index|
          hash[sheet.name.to_sym]["row_#{row_index + 1}".to_sym] = {}
          row.each_with_index do |val, index|
            if val.class == Spreadsheet::Formula
              if val.value.class == Spreadsheet::Excel::Error
                hash[sheet.name.to_sym]["row_#{row_index + 1}".to_sym]["column_#{index + 1}".to_sym] = FORMULA_ERROR
                next
              end
              hash[sheet.name.to_sym]["row_#{row_index + 1}".to_sym]["column_#{index + 1}".to_sym] = val.value
            else
              hash[sheet.name.to_sym]["row_#{row_index + 1}".to_sym]["column_#{index + 1}".to_sym] = val
            end
          end
        end
      end
    end

    def self.with_characters(book, hash)
      book.worksheets.each do |sheet|
        sheet.each_with_index do |row, row_index|
          hash[sheet.name.to_sym]["#{row_index + 1}".to_sym] = {}
          row.each_with_index do |val, index|
            if val.class == Spreadsheet::Formula
              if val.value.class == Spreadsheet::Excel::Error
                hash[sheet.name.to_sym]["#{row_index + 1}".to_sym]["#{Converter.to_characters(index)}".to_sym] = FORMULA_ERROR
                next
              end
              hash[sheet.name.to_sym]["#{row_index + 1}".to_sym]["#{Converter.to_characters(index)}".to_sym] = val.value
            else
              hash[sheet.name.to_sym]["#{row_index + 1}".to_sym]["#{Converter.to_characters(index)}".to_sym] = val
            end
          end
        end
      end
    end

    def self.kill(object)
      object = nil
      GC.start
    end

  end

end