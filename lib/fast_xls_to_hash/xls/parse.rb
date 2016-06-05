require 'spreadsheet'

module FastXlsToHash

  class ParseXls

    FORMULA_ERROR = 'Error'

    def self.to_hash(file)
      book = Spreadsheet.open(file)

      hash = Hash.new{ |hash, key| hash[key] = {} }
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

      kill(book)

      hash
    end

    private

    def kill(object)
      object = nil
      GC.start
    end

  end

end