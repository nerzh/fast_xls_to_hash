module FastXlsToHash
  class Converter
    CHAR = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']

    def self.to_characters(num)
      num += 1
      result = ''
      while(num > 0)
        num -= 1
        rest = num % 26
        num  = num / 26
        result += CHAR[rest]
      end
      result.reverse
    end
  end
end