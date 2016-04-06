# encoding: UTF-8

require "docx_replace/version"
require 'zip'

module DocxReplace
  class Doc
    attr_reader :document_content

    def initialize(path)
      @zip_file = Zip::File.new(path)
      read_docx_file
    end

    def replace(pattern, replacement, multiple_occurrences=false)
      replace = replacement.to_s
      if multiple_occurrences
        @document_content.force_encoding("UTF-8").gsub!(pattern, replace)
      else
        @document_content.force_encoding("UTF-8").sub!(pattern, replace)
      end
    end

    def matches(pattern)
      @document_content.scan(pattern).map{|match| match.first}
    end

    def unique_matches(pattern)
      matches(pattern)
    end

    alias_method :uniq_matches, :unique_matches


    def commit(new_path=nil)
      write_back_to_file(new_path)
    end

    private
    DOCUMENT_FILE_PATH = 'word/document.xml'

    def read_docx_file
      @document_content = @zip_file.read(DOCUMENT_FILE_PATH)
    end

    def write_back_to_file(new_path=nil)
      Zip::OutputStream.write_buffer do |zos|
        @zip_file.entries.each do |e|
          unless e.name == DOCUMENT_FILE_PATH
            zos.put_next_entry(e.name)
            zos.write e.get_input_stream.read
          end
        end

        zos.put_next_entry(DOCUMENT_FILE_PATH)
        zos.write @document_content
      end.string
    end
  end
end
