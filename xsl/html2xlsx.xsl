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
  
  <xsl:template match="/">
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <sheetPr filterMode="false">
        <pageSetUpPr fitToPage="false"/>
      </sheetPr>
      <dimension ref="A1:I48"/>
      <sheetViews>
        <sheetView tabSelected="1" topLeftCell="A1" zoomScaleNormal="100" workbookViewId="0">
          <selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"/>
        </sheetView>
      </sheetViews>
<!--      <sheetFormatPr baseColWidth="10" defaultColWidth="8.88671875" defaultRowHeight="150" />-->
      <!--<cols> 
        <xsl:call-template name="cols"/>
      </cols>-->
      <sheetData> 
        <xsl:apply-templates/>
      </sheetData>
    </worksheet>
  </xsl:template>
  
  <xsl:template match="*:title"/>
  
   <xsl:template match="*:html| *:body | *:head | *:div |*:a |*:p |*:span[ancestor::*:p] |*:i |*:b |*:sub| *:sup ">
    <xsl:apply-templates/>
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
    <xsl:apply-templates select="*:tr[2]" >
      <xsl:with-param name="previousRow" select="*:tr[1]" />
      <xsl:with-param name="rownum" select="$start-rownum + 1" />
    </xsl:apply-templates>
  </xsl:template>
  
  <xsl:template match="*:tr">
    <xsl:param name="previousRow" as="element()?" />
    <xsl:param name="rownum" as="xs:integer?" />
    <xsl:variable name="currentRow" select="." />
    <row r="{$rownum}" customFormat="true" ht="150" hidden="false" outlineLevel="0" collapsed="false">
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
    <c r="{concat($colnum,$rownum)}" t="s">
      <v>
        <xsl:apply-templates/>
      </v>
    </c>
  </xsl:template>
  
  
<!--   create shared strings from that?, new mode over all <v> values-->
  
  
  
</xsl:stylesheet>