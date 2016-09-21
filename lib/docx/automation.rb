require "cgi"
require "docx"
require "nokogiri"

require "docx/automation/field"
require "docx/automation/version"

module Docx
  module Automation
    extend self

    def render(template:, data:)
      docx = Docx::Document.new(template)

      docx.content_docs.each do |name, doc|
        table_loop_render(doc.root, data)
        Field.render(doc.root, data)
      end

      tmp_path = generate_tmp_path
      docx.save(tmp_path)
      tmp_path
    end

    def table_loop_render(element, data)
      element.css("w|tblCaption").select do |table_caption_el|
        caption = table_caption_el.attributes.fetch("val").value.strip
        next unless caption.start_with?("loop ")
        list_name = caption.split(/\s+/).last
        list = data.fetch(list_name)

        table_el = table_caption_el.parent.parent # TODO: Make this more robust
        header_row_el = table_el.css("w|tr")[0] # TODO: Make this more efficient (lookup once)
        content_row_el = table_el.css("w|tr")[1]

        list.each do |hash|
          new_content_row_el = content_row_el.dup
          Field.render(new_content_row_el, data.dup.merge(hash))
          content_row_el.add_previous_sibling(new_content_row_el)
        end

        content_row_el.remove # Drop template row
      end
    end

    # Generates file unnecessarily due to limitations in macOS version of mktemp
    def generate_tmp_path
      `mktemp /tmp/docx-auto.XXXXXXXX`.strip + ".docx"
    end
  end
end
