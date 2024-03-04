require 'roo/excelx/extractor'

module Roo
  class Excelx
    class Workbook < Excelx::Extractor
      class Label
        attr_reader :sheet, :row, :col, :name

        def initialize(name, sheet, row, col)
          @name = name
          @sheet = sheet
          @row = row.to_i
          @col = ::Roo::Utils.letter_to_number(col)
        end

        def key
          [@row, @col]
        end
      end

      def initialize(path)
        super
        fail ArgumentError, 'missing required workbook file' unless doc_exists?
      end

      def sheets
        doc.xpath('//sheet')
      end

      # aka labels
      def defined_names
        doc.xpath('//definedName').each_with_object({}) do |defined_name, hash|
          # "Sheet1!$C$5"
          sheet, coordinates = defined_name.text.split('!$', 2)
          next unless coordinates
          col, row = coordinates.split('$')
          name = defined_name['name']
          hash[name] = Label.new(name, sheet, row, col)
        end
      end

      def defined_names_in_book
        defined_name_objects_in_book = doc.xpath('//definedName').select do |defined_name_object|
          !defined_name_object["localSheetId"]
        end
        defined_name_objects_in_book.map { |defined_name_object| defined_name_object["name"] }
      end

      def defined_names_in_sheets
        defined_name_objects_in_sheets = doc.xpath('//definedName').select do |defined_name_object|
          defined_name_object["localSheetId"]
        end
        defined_name_objects_in_sheets.each_with_object({}) do |defined_name_object, result|
          local_sheet_id = defined_name_object["localSheetId"].to_i
          local_sheet_name = sheets[local_sheet_id]["name"]
          result[local_sheet_name] ||= []
          result[local_sheet_name] << defined_name_object["name"]
        end
      end

      def base_timestamp
        @base_timestamp ||= base_date.to_datetime.to_time.to_i
      end

      def base_date
        @base_date ||=
        begin
          # Default to 1900 (minus one day due to excel quirk) but use 1904 if
          # it's set in the Workbook's workbookPr
          # http://msdn.microsoft.com/en-us/library/ff530155(v=office.12).aspx
          result = Date.new(1899, 12, 30) # default
          doc.css('workbookPr[date1904]').each do |workbookPr|
            if workbookPr['date1904'] =~ /true|1/i
              result = Date.new(1904, 01, 01)
              break
            end
          end
          result
        end
      end
    end
  end
end
