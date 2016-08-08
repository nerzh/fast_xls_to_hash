require "fast_xls_to_hash/version"
require "fast_xls_to_hash/convert_number_to_xls_column/number_to_xls_column"
require "fast_xls_to_hash/xls/parse"

module FastXlsToHash
  @@xls_style = false

  def self.with_xls_style=(bool)
    @@xls_style = bool
  end

  def self.with_xls_style
    @@xls_style
  end
end
