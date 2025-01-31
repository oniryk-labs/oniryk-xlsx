export default function generateStyles() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
      <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <fonts count="1">
      <font>
          <sz val="11"/>
          <name val="Calibri"/>
          <family val="2"/>
      </font>
  </fonts>
  <fills count="1">
      <fill>
          <patternFill patternType="none"/>
      </fill>
  </fills>
  <borders count="1">
      <border>
          <left/><right/><top/><bottom/><diagonal/>
      </border>
  </borders>
  <cellStyleXfs count="1">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
      <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`;
}
