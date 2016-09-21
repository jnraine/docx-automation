require 'nokogiri'
require 'zip'

# Taken and modified from https://github.com/chrahunt/docx
module Docx
  # The Document class wraps around a docx file and provides methods to
  # interface with it.
  #
  #   # get a Docx::Document for a docx file in the local directory
  #   doc = Docx::Document.open("test.docx")
  #
  #   # get the text from the document
  #   puts doc.text
  #
  #   # do the same thing in a block
  #   Docx::Document.open("test.docx") do |d|
  #     puts d.text
  #   end
  class Document
    attr_reader :content_docs, :zip, :styles

    CONTENT_NAME_WHITELIST = [
      "word/document.xml",
      "word/footnotes.xml",
      "word/endnotes.xml",
      /word\/(footer|header)\d+.xml/,
    ]
    
    def initialize(path, &block)
      @replace = {}
      @zip = Zip::File.open(path)

      @content_docs = {}

      content_entries = @zip.select do |entry|
        CONTENT_NAME_WHITELIST.any? {|whitelist| entry.name.match(whitelist) }
      end

      content_entries.map do |content_entry|
        xml = @zip.read(content_entry.name)
        @content_docs[content_entry.name] = Nokogiri::XML(xml)
      end

      @styles_xml = @zip.read('word/styles.xml')
      @styles = Nokogiri::XML(@styles_xml)

      if block_given?
        yield self
        @zip.close
      end
    end

    # With no associated block, Docx::Document.open is a synonym for Docx::Document.new. If the optional code block is given, it will be passed the opened +docx+ file as an argument and the Docx::Document oject will automatically be closed when the block terminates. The values of the block will be returned from Docx::Document.open.
    # call-seq:
    #   open(filepath) => file
    #   open(filepath) {|file| block } => obj
    def self.open(path, &block)
      self.new(path, &block)
    end

    # call-seq:
    #   save(filepath) => void
    def save(path)
      update

      Zip::OutputStream.open(path) do |out|
        zip.each do |entry|
          out.put_next_entry(entry.name)

          if @replace[entry.name]
            out.write(@replace[entry.name])
          else
            out.write(zip.read(entry.name))
          end
        end
      end

      zip.close
    end

    alias_method :text, :to_s

    def replace_entry(entry_path, file_contents)
      @replace[entry_path] = file_contents
    end

    private

    def update
      content_docs.each do |name, doc|
        replace_entry name, doc.serialize(save_with: 0)
      end
    end
  end
end
