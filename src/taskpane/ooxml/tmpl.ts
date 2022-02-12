export const PackageTemplate = `
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        {{paragraphs}}
          <w:p/>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
`.replace(/\n\s*/g, "");

export const ParagraphTemplate = `
<w:p>
  <w:pPr>
    <w:spacing w:after="10" w:before="10" w:line="{{lineHeight}}" w:lineRule="auto"/>
    {{borderNode}}
    {{fillColorNode}}
    <w:ind w:left="{{indentation}}" w:hanging="{{indentation}}"/>
    <w:jc w:val="left"/>
  </w:pPr>
  {{texts}}
</w:p>
`.replace(/\n\s*/g, "");

export const FillColorNodeTemplate = `<w:shd w:fill="{{fillColor}}"/>`;

export const BorderTemplate = `
<w:pBdr>
  <w:top w:val="single" w:sz="4" w:space="1" w:color="auto" />
  <w:left w:val="single" w:sz="4" w:space="4" w:color="auto" />
  <w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto" />
  <w:right w:val="single" w:sz="4" w:space="4" w:color="auto" />
</w:pBdr>
`.replace(/\n\s*/g, "");

export const TextTemplate = `
<w:r>
  <w:rPr>
    <w:rFonts w:ascii="{{font}}"/>
    <w:sz w:val="{{fontSize}}"/>
    <w:color w:val="{{textColor}}"/>
  </w:rPr>
  <w:t xml:space="preserve">{{text}}</w:t>
</w:r>
`.replace(/\n\s*/g, "");