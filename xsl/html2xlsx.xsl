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
  
  <xsl:output method="xml" indent="yes" standalone="yes"/>
  
  <xsl:param name="th-style-from-row" select="0" as="xs:integer"/>
  <xsl:param name="td-style-from-row" select="0" as="xs:integer"/>
  
  <xsl:param name="keep-rows" select="0"  as="xs:integer"/>
  <xsl:param name="static-rows"  as="xs:string" select="'no'" />
  <xsl:param name="use-html-th" select="'yes'"/>
  
  <xsl:variable name="template" select="collection()[/*:worksheet]"/>
  <xsl:variable name="html" select="collection()[/*:html]"/>
  <xsl:variable name="shared-strings-root" select="collection()[/*:sst]" as="document-node()?"/>
  
  <xsl:variable name="is-html-th-used" select="$use-html-th = 'yes'" as="xs:boolean"/>
  <xsl:variable name="are-rows-static" select="$static-rows = 'yes'" as="xs:boolean"/>
  
  <xsl:variable name="alphabet-sequence" select="('A','B','C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P',
                                                  'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z')"/>
  
  <xsl:key name="string-by-si" match="*:si" use="count(preceding-sibling::*:si)"/>
  
  <xsl:template match="/">
    <xsl:copy>
      <xsl:apply-templates/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="@* | node()" mode="#default relation shared-strings subst gen-shared-strings">
    <xsl:copy inherit-namespaces="no">
      <xsl:apply-templates select="@*, node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="autoFilter">
    <xsl:copy>
      <xsl:attribute name="ref" select="concat('A1:', replace($template//*:row[position() = $th-style-from-row]/*:c[last()]/@*:r, '[0-9]+$', count($html//*:tr) cast as xs:string))"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:sheetData">
    <xsl:copy>
      <xsl:if test="$keep-rows &gt; 0 and $are-rows-static">
        <xsl:apply-templates select="$template//*:row[position() &lt;= $keep-rows]"/>
      </xsl:if>
      <xsl:apply-templates select="$html"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:thead">
    <xsl:if test="$is-html-th-used">
      <xsl:message select="'th-style-from-row: ', $th-style-from-row"/>
      <xsl:apply-templates select="*:tr"/>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="*:tbody">
    <xsl:message select="'td-style-from-row: ', $td-style-from-row "/>
    <xsl:apply-templates select="*:tr"/>
  </xsl:template>
  
  <xsl:template match="*:tr">
    <xsl:variable name="tr_count" select="count(if ($is-html-th-used) then preceding::*:tr else preceding::*:tr[not(ancestor::*:thead)])+1" as="xs:integer"/>
    <xsl:variable name="self" select="."/>
    <xsl:variable name="row-template" as="element()?">
      <xsl:choose>
        <xsl:when test="$td-style-from-row &gt; 0">
          <xsl:sequence select="$template//*:row[position() = $td-style-from-row]"/>
        </xsl:when>
        <xsl:when test="not($are-rows-static) and ($td-style-from-row eq 0)">
          <xsl:sequence select="$template//*:row[position() = $tr_count]"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="row-num" as="xs:integer" select="if ($are-rows-static) then $tr_count +$keep-rows else $tr_count"/>
    <xsl:message select="'keep-rows: ',$keep-rows, 'row-num: ', $row-num"/>
    <xsl:element name="row">
      <xsl:apply-templates select="$row-template/@*"/>
      <xsl:attribute name="r" select="$row-num"/>
      <xsl:apply-templates select="node()">
        <xsl:with-param name="row-template" select="$row-template" tunnel="yes"/>
        <xsl:with-param name="row-num" as="xs:integer" select="$row-num" tunnel="yes"/>
      </xsl:apply-templates>
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="*:th | *:td">
    <xsl:param name="row-template" as="element()?" tunnel="yes"/>
    <xsl:param name="row-num" as="xs:integer" tunnel="yes"/>
    <xsl:variable name="pos" as="xs:integer" select="count(preceding-sibling::*)+1"/>
    <xsl:variable name="text" as="xs:string" select="string-join(descendant::text(),'')"/>
    <xsl:element name="c">
      <xsl:for-each select="$row-template/*:c[$pos]">
        <xsl:apply-templates select="@* except (@*:r | @*:t)"/>
        <xsl:attribute name="r" select="replace(@*:r, '^([A-Z]+)([0-9]+)$', concat('$1', $row-num))"/>
        <xsl:attribute name="t">
          <xsl:choose>
            <xsl:when test="matches($text, '^[\-|\+\*/:]?\d+$')">n</xsl:when>
            <xsl:when test="$text eq '-' or *:f" >str</xsl:when>
            <xsl:otherwise>s</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:for-each select="*:f">
          <xsl:copy inherit-namespaces="no">
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
        <xsl:element name="v" inherit-namespaces="no">
          <xsl:apply-templates select="node()"/>
        </xsl:element>
      </xsl:if>
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="*:extLst|*:dataValidations"/>
  
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
  
  <xsl:template match="*:a[@href]" priority="2">
    <xsl:variable name="id" select="concat('rId',replace(generate-id(), '[a-z]', ''))"/>
    <xsl:element name="hyperlink" inherit-namespaces="no">
      <xsl:attribute name="id" select="$id" namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
      <xsl:attribute name="display" select="string(.)"/>
        <xsl:element name="Relationship" namespace="http://schemas.openxmlformats.org/package/2006/relationships" exclude-result-prefixes="#all">
        <xsl:attribute name="Id" select="$id"/>
        <xsl:attribute name="Type" select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'" />
        <xsl:attribute name="Target" select="@href" />
        <xsl:attribute name="TargetMode" select="'External'"/>
        </xsl:element>
    </xsl:element>
  </xsl:template>
  
<!--  MODE relation-->
  
  <xsl:template match="*:hyperlink" mode="relation">
    <xsl:value-of select="@display"/>
    <xsl:result-document href="{generate-id()}">
      <xsl:sequence select="*:Relationship"/>
    </xsl:result-document>
  </xsl:template>
  
  <xsl:template match="*:worksheet/*:hyperlinks" mode="relation">
    <xsl:copy inherit-namespaces="no">
      <xsl:apply-templates select="@*, node()" mode="relation"/>
      <xsl:apply-templates select="//*:hyperlink" mode="hyperlinks"/>
    </xsl:copy>
  </xsl:template>
  <xsl:template match="*:worksheet/*:hyperlinks" mode="relation"/>
  
  <xsl:template match="*:worksheet[not(*:hyperlinks) and exists(//*:hyperlink)]/*:pageMargins" mode="relation">
    <hyperlinks>
      <xsl:apply-templates select="//*:hyperlink" mode="hyperlinks"/>
    </hyperlinks>
    <xsl:copy>
      <xsl:apply-templates select="@*, node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:hyperlink" mode="hyperlinks">
    <xsl:copy inherit-namespaces="no">
      <xsl:attribute name="ref" select="ancestor::*:c[1]/@r"/>
      <xsl:if test="not(ancestor::*:c[1]/@r)">
        <xsl:attribute name="ref">
          <xsl:variable name="count_sibs" select="count(ancestor::*:c[1]/preceding-sibling::*:c) + 1"/>
          <xsl:call-template name="compute-col-chars">
            <xsl:with-param name="position" select="$count_sibs"/>
          </xsl:call-template>
          <xsl:value-of select="ancestor::*:row[@r][1]/@r"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:apply-templates select="@*" mode="relation"/>
    </xsl:copy>
  </xsl:template>
  
<!--  MODE shared-strings-->
  
  <xsl:template match="/" mode="shared-strings">
    <xsl:variable name="resolved-template-strings" as="element()?">
      <xsl:apply-templates select="*" mode="shared-strings"/>
    </xsl:variable>
    <xsl:apply-templates select="$resolved-template-strings" mode="subst"/>
    <!--<xsl:result-document href="{'bogo.xml'}">
      <xsl:document>
        <sst count="{count($resolved-template-strings//*:v)}" uniqueCount="{count($resolved-template-strings//*:v)}" >
          <xsl:apply-templates select="$resolved-template-strings//*:v" mode="gen-shared-strings"/>
        </sst>
      </xsl:document>
    </xsl:result-document>-->
  </xsl:template>
  
  <xsl:template match="*:row[@r &lt;= $keep-rows and $are-rows-static][*:c[@t = 's']]//*:v" mode="shared-strings">
      <xsl:apply-templates select="key('string-by-si', number(text()), $shared-strings-root)" mode="#current"/>
  </xsl:template>
  
  <xsl:template match="*:si" mode="shared-strings">
    <is>
        <xsl:apply-templates select="node()" mode="#current"/>
    </is>
  </xsl:template>
  
 <!-- <xsl:template match="*:v[*:si]" mode="gen-shared-strings">
      <xsl:apply-templates select="node()" mode="#current"/>
  </xsl:template>
  
  <xsl:template match="*:v/*:si" mode="gen-shared-strings">
    <xsl:copy>
      <xsl:apply-templates select="node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:v[not(*:si)]" mode="gen-shared-strings">
    <si>
      <t>
        <xsl:apply-templates select="node()" mode="#current"/>
      </t>
    </si>
  </xsl:template>-->
  
  <xsl:template match="*:c[@t = 's' or not(@t)][*:v]" mode="subst">
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:if test="not(@r)">
        <xsl:call-template name="gen-r-attr"/>
      </xsl:if>
      <xsl:attribute name="t" select="'inlineStr'"/>
      <xsl:apply-templates select="node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:c[@t = 's' or not(@t)]/*:v" mode="subst">
      <is>
        <t>
          <xsl:apply-templates select="node()" mode="#current"/>
        </t>
      </is>
  </xsl:template>
  
  <xsl:template name="compute-col-chars">
    <xsl:param name="position" as="xs:decimal"/>
    <xsl:param name="alpha-places" select="1" as="xs:decimal"/>
    <xsl:param name="last-alpha-char" as="xs:string?">
    </xsl:param>
    <xsl:variable name="alpha-pl" select="floor($position div 26)" as="xs:decimal"/>
    <xsl:variable name="pl" select="if (($position - $alpha-pl*26) = 0) then 26 else $position - ($alpha-pl*26)"/>
    <xsl:choose>
      <xsl:when test="$position &lt;= 26">
        <xsl:sequence select="string-join(($last-alpha-char, $alphabet-sequence[$pl]), '')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="compute-col-chars">
          <xsl:with-param name="position" select="$pl"/>
          <xsl:with-param name="last-alpha-char" select="string-join(($last-alpha-char,$alphabet-sequence[if ($pl = 26) then ($alpha-pl - 1) else $alpha-pl]), '')"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="*:c[empty(node())][not(@r)]" mode="subst">
    <xsl:variable name="position" select="count(preceding-sibling::*:c)+1"/>
    <xsl:variable name="col_chars">
      <xsl:call-template name="compute-col-chars">
        <xsl:with-param name="position" select="$position"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:copy>
      <xsl:call-template name="gen-r-attr"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:c[empty(node())]/@t" priority="3" mode="subst"/>
  
  <xsl:template match="*:c/@t[. = 's']" priority="2" mode="subst">
    <xsl:attribute name="t" select="'inlineStr'"/>
  </xsl:template>
  
  <!--<xsl:template match="*:c/@s" mode="subst"/>-->
  
  <xsl:template match="*:hyperlink" mode="subst">
    <xsl:copy inherit-namespaces="no">
      <xsl:apply-templates select="@*,node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template name="gen-r-attr">
    <xsl:variable name="count_sibs" select="count(preceding-sibling::*:c)+1"/>
    <xsl:variable name="col_char">
      <xsl:call-template name="compute-col-chars">
        <xsl:with-param name="position" select="$count_sibs"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:attribute name="r" select="concat($col_char, ../@r)"/>
  </xsl:template>
  
  <!--<xsl:template match="row[@r &lt;= $keep-rows]//v" mode="shared-strings">
    <xs:copy>
      <xsl:apply-templates select="key('string-by-v', ., $shared-strings-root)" mode="#current"/>
    </xs:copy>
  </xsl:template>-->
  
</xsl:stylesheet>