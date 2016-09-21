require 'spec_helper'

describe Docx::Automation do
  def doc_path(name)
    path = File.expand_path("../support/#{name}", File.dirname(__FILE__))
    raise "#{path} does not exist" unless File.exist?(path)
    path
  end

  it 'has a version number' do
    expect(Docx::Automation::VERSION).not_to be nil
  end

  # it 'does something useful' do
  #   data = {"first_name" => "Joe"}
  #   rendered_file = Docx::Automation.render(template: doc_path("letter.docx"), data: data)
  #   `open #{rendered_file}`
  # end

  let(:support_path) { Pathname(__FILE__).parent.parent.join("support") }

  xit "generates things" do
    tmp_doc = Docx::Automation.render(
      template: doc_path("id.docx"),
      data: {
        "hello" => "Foo bar baz",
        "quantity" => Time.now.to_s,
        "description" => "This is an item"
      }
    )

    `open #{tmp_doc}`
  end

  xit "does cool things with tables" do
    tmp_doc = Docx::Automation.render(
      template: doc_path("table.docx"),
      data: {
        "line_items" => [
          {"description" => "Discovery for Lenny", "quantity" => 12.5, "id" => rand(1000)},
          {"description" => "Filing evidence", "quantity" => 3, "id" => rand(1000)},
          {"description" => "Lunch with Phil", "quantity" => 1, "id" => rand(1000)},
        ],
      }
    )

    `open #{tmp_doc}`
  end

  it "handles nested variables" do
    tmp_doc = Docx::Automation.render(
      template: doc_path("dot-notation.docx"),
      data: {
        "firm" => {
          "name" => "Cooper Sterling Price",
          "address" => {
            "street" => "707 East 20th Ave",
            "city" => "Vancouver",
            "province" => "British Columbia",
            "postal_code" => "V5V 0B3"
          }
        },
        "invoice" => {
          "date" => "2016-09-20",
          "id" => "000867"
        }
      }
    )

    `open #{tmp_doc}`
  end
end
