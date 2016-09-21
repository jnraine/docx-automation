require 'spec_helper'


describe Docx::Automation::Field do
  let(:example_doc) do

    example_xml = <<-XML
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">
<w:body>
  <w:p w14:paraId="312F0D1F" w14:textId="77777777" w:rsidR="00892973" w:rsidRDefault="00892973">
    <w:r>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:color w:val="FFFF00"/>
      </w:rPr>
      <w:instrText>firm.address.city</w:instrText>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
    <w:r>
      <w:t xml:space="preserve">, </w:t>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <w:instrText>firm.address.province</w:instrText>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </w:p>
</w:body>
</w:document>
XML

    example_doc = Nokogiri::XML(example_xml)
  end

  it "parses shit" do
    fields = Docx::Automation::Field.from_element(example_doc.root)

    expect(fields.length).to eq(2)
    expect(fields.first.name).to eq("firm.address.city")
    expect(fields.first.runs.length).to eq(3)

    expect(fields.last.name).to eq("firm.address.province")
    expect(fields.last.runs.length).to eq(3)
  end

  it "replaces shit" do
    data = {
      "firm" => {
        "address" => {
          "city" => "Vancouver",
          "province" => "British Columbia",
        }
      }
    }

    fields = Docx::Automation::Field.from_element(example_doc.root)
    fields.map {|field| field.render(data) }

    puts example_doc
  end
end
