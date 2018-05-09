<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc"
  xmlns:c="http://www.w3.org/ns/xproc-step" 
  xmlns:cx="http://xmlcalabash.com/ns/extensions"
  xmlns:pxp="http://exproc.org/proposed/steps"
  xmlns:cxf="http://xmlcalabash.com/ns/extensions/fileutils"
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:tr="http://transpect.io"
  version="1.0" 
  name="overwrite-files"
  type="tr:overwrite-files">
  
  
  <p:input port="source">
    <p:documentation>replacement</p:documentation>
  </p:input>
  
  
  <p:option name="file" required="true"/>
  <p:option name="debug" select="'yes'"/>
  <p:option name="debug-dir-uri" select="'debug'"/>
  
  <p:import href="http://transpect.io/xproc-util/file-uri/xpl/file-uri.xpl"/>
  <p:import href="http://transpect.io/xproc-util/store-debug/xpl/store-debug.xpl"/>
  <p:import href="http://transpect.io/xproc-util/xslt-mode/xpl/xslt-mode.xpl"/>
  
  
  <tr:file-uri name="file-uri">
    <p:with-option name="filename" select="$file"/>
  </tr:file-uri>
  
  <tr:store-debug pipeline-step="excel/file-uri">
    <p:with-option name="active" select="$debug"/>
    <p:with-option name="base-uri" select="$debug-dir-uri"/>
  </tr:store-debug>
  
  <p:sink/>
      
  <p:load name="load">
    <p:with-option name="href" select="/c:result/@local-href">
      <p:pipe port="result" step="file-uri"/>
    </p:with-option>
  </p:load>
  
  <p:replace match="/">
    <p:input port="replacement">
      <p:pipe port="source" step="overwrite-files"/>
    </p:input>
  </p:replace>
  
  <p:store name="store">
    <p:with-option name="href" select="/c:result/@local-href">
      <p:pipe port="result" step="file-uri"/>
    </p:with-option>
  </p:store>
  
</p:declare-step>