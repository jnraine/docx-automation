module Docx
  module Automation
    class Field
      DoubleRenderError = Class.new(StandardError)

      # Right now, this edits the element passed in. Should it dup and return a new, modified element?
      def self.render(element, data)
        from_element(element).each {|field| field.render(data) }
      end

      def self.from_element(element)
        element.css("w|fldChar[w|fldCharType='begin']").map do |begin_element|
          Field.new(begin_element)
        end
      end

      attr_reader :begin_element, :end_element, :name, :runs

      def initialize(begin_element)
        @begin_element = begin_element
        @name = ""
        @end_element = nil
        @runs = []

        next_el =  begin_element.parent
        @runs << next_el
        loop do
          next_el = next_el.next_sibling
          break if next_el.nil?

          # TODO: Should all runs after the begin be considered part of the field?)
          if valid_run?(next_el)
            @runs << next_el
            @name += next_el.css("w|instrText").map(&:text).join
            if contains_field_end?(next_el)
              @end_element = next_el.css("w|fldChar[w|fldCharType='end']").first
              break
            end
          end
        end
      end

      def valid_run?(element)
        element.name == "r" && element.css("w|fldChar,w|instrText")
      end

      def contains_field_end?(element)
        element.css("w|fldChar[w|fldCharType='end']").any?
      end

      def render(data)
        raise DoubleRenderError.new("This field has already been rendered") if rendered?

        value = dig(data, *name.split("."))

        # Grab formatting elements then remove duplicates
        style_children = runs.flat_map do |child|
          child.css("w|rPr").flat_map(&:children)
        end.each_with_object({}) {|child, hash| hash[child.name] = child }.values

        style_el = Nokogiri::XML::Element.new("w:rPr", begin_element.document)
        style_children.map {|style_child| style_el << style_child }

        text_el = Nokogiri::XML::Element.new("w:t", begin_element.document)
        text_el.inner_html = CGI.escape_html(value.to_s)

        new_value_el = Nokogiri::XML::Element.new("w:r", begin_element.document)
        new_value_el << style_el
        new_value_el << text_el

        # Insert new element after last run, then remove field runs, then mark as rendered
        runs.last.add_next_sibling(new_value_el)
        runs.map(&:remove)
        @rendered = true
      end

      def rendered?
        @rendered
      end

      # Hey, we're not on Ruby 2.3 yet!
      private def dig(hash, *keys)
        last_hash = hash

        keys.each do |key|
          break if last_hash.nil? || !last_hash.respond_to?(:[])
          last_hash = last_hash[key]
        end

        last_hash
      end
    end
  end
end
