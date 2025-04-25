package docx

const defaultStylesXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="240" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
       <w:spacing w:after="160" w:line="240" w:lineRule="auto"/>
    </w:pPr>
     <w:rPr>
        <w:sz w:val="22"/>
     </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="DefaultParagraphFont" w:default="1">
     <w:name w:val="Default Paragraph Font"/>
     <w:uiPriority w:val="1"/>
     <w:semiHidden/>
     <w:unhideWhenUsed/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="Heading1Char"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="240" w:after="0"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
      <w:b/>
      <w:bCs/>
      <w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF"/>
      <w:sz w:val="32"/>
      <w:szCs w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Heading1Char">
    <w:name w:val="Heading 1 Char"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:link w:val="Heading1"/>
    <w:uiPriority w:val="9"/>
    <w:rPr>
      <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
      <w:b/>
      <w:bCs/>
      <w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF"/>
      <w:sz w:val="32"/>
      <w:szCs w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="Heading2Char"/>
    <w:uiPriority w:val="9"/>
    <w:unhideWhenUsed/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="40" w:after="0"/>
      <w:outlineLvl w:val="1"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
      <w:b/>
      <w:bCs/>
      <w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF"/>
      <w:sz w:val="26"/>
      <w:szCs w:val="26"/>
    </w:rPr>
  </w:style>
   <w:style w:type="character" w:styleId="Heading2Char">
    <w:name w:val="Heading 2 Char"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:link w:val="Heading2"/>
    <w:uiPriority w:val="9"/>
    <w:rPr>
      <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
      <w:b/>
      <w:bCs/>
      <w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF"/>
      <w:sz w:val="26"/>
      <w:szCs w:val="26"/>
    </w:rPr>
  </w:style>
   <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="Heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="Heading3Char"/>
    <w:uiPriority w:val="9"/>
    <w:unhideWhenUsed/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="40" w:after="0"/>
      <w:outlineLvl w:val="2"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
      <w:b/>
      <w:bCs/>
      <w:color w:val="1F4D78" w:themeColor="accent1" w:themeShade="7F"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Heading3Char">
     <w:name w:val="Heading 3 Char"/>
     <w:basedOn w:val="DefaultParagraphFont"/>
     <w:link w:val="Heading3"/>
     <w:uiPriority w:val="9"/>
     <w:rPr>
        <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
        <w:b/>
        <w:bCs/>
        <w:color w:val="1F4D78" w:themeColor="accent1" w:themeShade="7F"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
     </w:rPr>
  </w:style>
   <w:style w:type="paragraph" w:styleId="Heading4">
    <w:name w:val="Heading 4"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="Heading4Char"/>
    <w:uiPriority w:val="9"/>
    <w:unhideWhenUsed/>
    <w:qFormat/>
    <w:pPr>
       <w:keepNext/>
       <w:keepLines/>
       <w:spacing w:before="40" w:after="0"/>
       <w:outlineLvl w:val="3"/>
    </w:pPr>
    <w:rPr>
       <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
       <w:i/>
       <w:iCs/>
       <w:color w:val="1F4D78" w:themeColor="accent1" w:themeShade="7F"/>
       <w:sz w:val="22"/>
       <w:szCs w:val="22"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Heading4Char">
     <w:name w:val="Heading 4 Char"/>
     <w:basedOn w:val="DefaultParagraphFont"/>
     <w:link w:val="Heading4"/>
     <w:uiPriority w:val="9"/>
     <w:rPr>
        <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
        <w:i/>
        <w:iCs/>
        <w:color w:val="1F4D78" w:themeColor="accent1" w:themeShade="7F"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
     </w:rPr>
  </w:style>
</w:styles>
`
