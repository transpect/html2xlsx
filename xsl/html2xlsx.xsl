<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:html="http://www.w3.org/1999/xhtml"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:tr="http://transpect.io"
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  version="2.0" xpath-default-namespace="http://www.w3.org/1999/xhtml" exclude-result-prefixes="#all">
  
  <xsl:output method="xml" indent="yes"/>
  
  <xsl:param name="th-template-row" as="xs:integer"/>
  <xsl:param name="td-template-row" as="xs:integer"/>
  
  <xsl:variable name="template" select="collection()[/*:worksheet]"/>
  <xsl:variable name="html" select="collection()[/*:html]"/>
  <!-- copy the first header rows from template, 
    if you don't want anything to be copied leave empty -->
  <xsl:variable name="keep-firstrows-from-worksheet" select="''" as="xs:integer"/>
  <xsl:variable name="use-html-th" select="false()" />
  
  <xsl:template match="/">
    <xsl:copy>
      <xsl:apply-templates/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="@* | node()">
    <xsl:copy>
      <xsl:apply-templates select="@*, node()"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="autoFilter">
    <xsl:copy>
      <xsl:attribute name="ref" select="concat('A1:', replace($template//*:row[position() = $th-template-row]/*:c[last()]/@*:r, '[0-9]+$', count($html//*:tr) cast as xs:string))"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:sheetData">
    <xsl:copy>
      <xsl:if test="$keep-firstrows-from-worksheet">
        <xsl:apply-templates select="$template//*:row[position() &lt;= $keep-firstrows-from-worksheet]"/>
      </xsl:if>
      <xsl:apply-templates select="$html"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:thead">
    <xsl:choose>
      <xsl:when test="$use-html-th">
        <xsl:message select="'th-template-row: ', $th-template-row "></xsl:message>
        <xsl:apply-templates select="*:tr">
          <xsl:with-param name="row-template" as="element()" select="$template//*:row[position() = $th-template-row]" tunnel="yes"/>
        </xsl:apply-templates>    
      </xsl:when>
      <xsl:otherwise>
        <xsl:copy-of select="$template//*:row[position() = $th-template-row]"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="*:tbody">
    <xsl:message select="'td-template-row: ', $td-template-row "></xsl:message>
    <xsl:apply-templates select="*:tr">
      <xsl:with-param name="row-template" as="element()" select="$template//*:row[position() = $td-template-row]" tunnel="yes"/>
    </xsl:apply-templates>    
  </xsl:template>
  
  <xsl:template match="*:tr">
    <xsl:param name="row-template" as="element()" tunnel="yes"/>
    <xsl:variable name="row-num" as="xs:integer" select="count(preceding::*:tr)+($keep-firstrows-from-worksheet,1)[1]"/>
    <xsl:message select="'keep-firstrows-from-worksheet: ',$keep-firstrows-from-worksheet, 'row-num: ', $row-num "></xsl:message>
    <xsl:element name="row">
      <xsl:apply-templates select="$row-template/@*"/>
      <xsl:attribute name="r" select="$row-num"/>
      <xsl:apply-templates select="node()">
        <xsl:with-param name="row-num" as="xs:integer" select="$row-num" tunnel="yes"/>
      </xsl:apply-templates>
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="*:th | *:td">
    <xsl:param name="row-template" as="element()" tunnel="yes"/>
    <xsl:param name="row-num" as="xs:integer" tunnel="yes"/>
    <xsl:variable name="pos" as="xs:integer" select="count(preceding-sibling::*)+1"/>
    <xsl:variable name="text" as="xs:string" select="string-join(descendant::text(),'')"/>
    <xsl:element name="c">
      <xsl:for-each select="$row-template/*:c[$pos]">
        <xsl:apply-templates select="@* except (@*:r | @*:t)"/>
        <xsl:attribute name="r" select="replace(@*:r, '^([A-Z]+)([0-9]+)$', concat('$1', $row-num))"/>
        <xsl:attribute name="t">
          <xsl:choose>
            <xsl:when test="matches($text, '^(\-|\+)?\d+$')">n</xsl:when>
            <xsl:when test="$text eq '-'">str</xsl:when>
            <xsl:otherwise>s</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:for-each select="*:f">
          <xsl:copy>
            <xsl:analyze-string select="." regex="([^A-Z][A-Z]+)([0-9]+)([^0-9])">
              <xsl:matching-substring>
                <xsl:value-of select="regex-group(1)"/>                
                <xsl:value-of select="if (regex-group(2) eq $row-template/@*:r) then $row-num else regex-group(2)"/>
                <xsl:value-of select="regex-group(3)"/>
              </xsl:matching-substring>
              <xsl:non-matching-substring>
                <xsl:value-of select="."/>
              </xsl:non-matching-substring>
            </xsl:analyze-string>
          </xsl:copy>
        </xsl:for-each>
      </xsl:for-each>
      <xsl:if test="$text[normalize-space()]">
        <xsl:element name="v">
          <xsl:apply-templates select="node()"/>
        </xsl:element>
      </xsl:if>
    </xsl:element>
  </xsl:template>
  
<!--  <xsl:template match="*:extLst"/>-->
  
  <!--<xsl:template match="*:cols">
    <xsl:copy>
      <xsl:apply-templates select="@*, node()"/>
    </xsl:copy>
  </xsl:template>-->
  
  <!--<xsl:function name="tr:col-to-num">
    <xsl:param name="col"/>
    <xsl:variable name="int" select="for $i in (1 to string-length($col)) return (string-to-codepoints(substring($col,$i,1))-64)"/>
<!-\-    <xsl:message select="$col, $int[1], count($int)"/>-\->
    <xsl:choose>
      <xsl:when test="count($int)=1">
        <xsl:sequence select="$int[1]"/>
      </xsl:when>
      <xsl:when test="count($int)=2">
        <xsl:sequence select="$int[1]*26+$int[2]"/>
      </xsl:when>
     <xsl:when test="count($int)=3">
        <xsl:sequence select="$int[1]*676+$int[2]*26+$int[3]"/>
      </xsl:when>
    </xsl:choose>
  </xsl:function>
  
  <xsl:template match="*:row[position()&lt;= count($html//*:tr)]">
    <xsl:copy>
      <xsl:message select="'process html row number ', string(@r), ' of ', count($html//*:tr),' rows'"/>
      <xsl:apply-templates select="@*, node()"/>
    </xsl:copy>
  </xsl:template>

  -->
  
  <xsl:template match="*:title| *:meta| *:style"/>
  
  <xsl:template match="*:html| *:body | *:head | *:div |*:a |*:p |*:span[ancestor::*:p] |*:i |*:b |*:sub| *:sup | *:table">
    <xsl:apply-templates/>
  </xsl:template>  
  
  <xsl:template match="*:br">
    <xsl:text>&#xa;</xsl:text>
  </xsl:template>
  
  
</xsl:stylesheet>