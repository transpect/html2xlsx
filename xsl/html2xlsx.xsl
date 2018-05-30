<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:html="http://www.w3.org/1999/xhtml"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  version="2.0" xpath-default-namespace="http://www.w3.org/1999/xhtml" exclude-result-prefixes="#all">
  
  <xsl:output method="xml" indent="yes"/>
  
<!--  <xsl:variable name="template" select="collection()[/*:worksheet]"/>-->
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
  
  <xsl:template match="*:sheetData">
    <xsl:copy>
      <xsl:apply-templates select="$html"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:extLst"/>
  
  <xsl:template match="*:cols">
    <xsl:copy>
<!--      <xsl:message select="."/>-->
      <xsl:apply-templates select="@*, node()"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:title| *:meta| *:style"/>
  
   <xsl:template match="*:html| *:body | *:head | *:div |*:a |*:p |*:span[ancestor::*:p] |*:i |*:b |*:sub| *:sup ">
    <xsl:apply-templates/>
  </xsl:template>  
  
  <xsl:template match="*:br">
    <xsl:text>&#xa;</xsl:text>
  </xsl:template>
  

  <xsl:template match="*[*:tr]">
    <xsl:variable name="start-rownum" select="count(preceding-sibling::*[*:tr]/*:tr) + 1"/>
    <xsl:for-each select="*:tr[1]">
      <row r="{$start-rownum}" customFormat="true" ht="150" hidden="false" outlineLevel="0" collapsed="false">
        <xsl:apply-templates select="*" >
          <xsl:with-param name="rownum" select="$start-rownum" />
        </xsl:apply-templates>
      </row>
    </xsl:for-each>
    <xsl:apply-templates select="descendant::*:tr[2]" >
      <xsl:with-param name="previousRow" select="*:tr[1]" />
      <xsl:with-param name="rownum" select="$start-rownum + 1" />
    </xsl:apply-templates>
  </xsl:template>
  
  <xsl:template match="*:tr">
    <xsl:param name="previousRow" as="element()?" />
    <xsl:param name="rownum" as="xs:integer?" />
    <xsl:variable name="currentRow" select="." />
    <row r="{$rownum}" customFormat="true" ht="150" hidden="false" outlineLevel="0" collapsed="false">
    <!-- get these attributes also from template, style info,  customFormat="1" aso -->
      <xsl:apply-templates select="*" >
        <xsl:with-param name="rownum" select="$rownum" />
      </xsl:apply-templates>
    </row>
    
    <xsl:variable name="newRow" as="element()">
      <xsl:copy>
        <xsl:copy-of select="$currentRow/@*" />
        <xsl:apply-templates/>
      </xsl:copy>
    </xsl:variable>
    
    <xsl:apply-templates select="following-sibling::*:tr[1]">
      <xsl:with-param name="previousRow" select="$newRow" />
      <xsl:with-param name="rownum" select="$rownum + 1" />
    </xsl:apply-templates>
    
  </xsl:template>
  
  
  <xsl:template match="*:td | *:th" >
    <xsl:param name="rownum" as="xs:integer?" />
    <xsl:variable name="colnum">
      <xsl:number value="position()" format="A"/>
    </xsl:variable>
    <c r="{concat($colnum,$rownum)}">
      <xsl:choose>
        <xsl:when test="matches(string-join(text(),''),'^(\-|\+)?\d+$') ">
            <xsl:attribute name="t" select="'n'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="t" select="'s'"/>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:attribute name="s" select="'139'"/>
      <v>
        <xsl:apply-templates/>
      </v>
    </c>
  </xsl:template>
  
  
<!--   create shared strings from that?, new mode over all <v> values-->
  
  
  
</xsl:stylesheet>