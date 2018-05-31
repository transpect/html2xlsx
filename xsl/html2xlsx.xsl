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
  
  <xsl:variable name="template" select="collection()[/*:worksheet]"/>
  <xsl:variable name="html" select="collection()[/*:html]"/>
  
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
  
  <!--<xsl:template match="*:sheetData">
    <xsl:copy>
      <xsl:apply-templates  />
    </xsl:copy>
  </xsl:template>-->
  
  <xsl:template match="*:extLst"/>
  
  <!--<xsl:template match="*:cols">
    <xsl:copy>
      <xsl:apply-templates select="@*, node()"/>
    </xsl:copy>
  </xsl:template>-->
  
  <xsl:function name="tr:col-to-num">
    <xsl:param name="col"/>
    <xsl:variable name="int" select="for $i in (1 to string-length($col)) return (string-to-codepoints(substring($col,$i,1))-64)"/>
<!--    <xsl:message select="$col, $int[1], count($int)"/>-->
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

  <xsl:template match="*:row[position() &gt; count($html//*:tr)+1]"/>

  <xsl:template match="*:c[position() &lt;= max($html//*:tr/count(*:td))]">
    <xsl:variable name="html-tr-position" select="number(replace(@r,'[A-Z]+',''))"/>
    <xsl:variable name="html-td-position" select="tr:col-to-num(replace(@r,'\d+',''))"/>
    <xsl:variable name="html-cell" select="//$html//(*:td|*:th)[parent::*:tr[count(preceding::*:tr)+1=$html-tr-position]]
                                            [position()=$html-td-position]"/>
    <xsl:message select="$html-cell"></xsl:message>
    <xsl:copy>
      <!-- read this styleinfomation from css or class? -->
      <xsl:choose>
        <xsl:when test="matches(string-join($html-cell/text(),''),'^(\-|\+)?\d+$') ">
          <xsl:attribute name="t" select="'n'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="t" select="'s'"/>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:apply-templates select="@* except @t"/>
      <!-- <xsl:message select="'row: ',$html-tr-position, 'cell: ', $html-td-position"></xsl:message>-->
      
      <v>
        <xsl:apply-templates select="$html-cell"/>
      </v>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:title| *:meta| *:style"/>
  
   <xsl:template match="*:html| *:body | *:head | *:div |*:a |*:p |*:span[ancestor::*:p] |*:i |*:b |*:sub| *:sup | *:table| *:td |*:th ">
    <xsl:apply-templates/>
  </xsl:template>  
  
  <xsl:template match="*:br">
    <xsl:text>&#xa;</xsl:text>
  </xsl:template>
  
  
</xsl:stylesheet>