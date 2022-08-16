var _, documentTemplate, fs, utils;

fs = require('fs');
var _ = require('underscore');

documentTemplate = _.template(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:ns6="http://schemas.openxmlformats.org/schemaLibrary/2006/main"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:ns8="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"
  xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:ns11="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
  xmlns:ns13="urn:schemas-microsoft-com:office:excel"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:ns17="urn:schemas-microsoft-com:office:powerpoint"
  xmlns:odx="http://opendope.org/xpaths"
  xmlns:odc="http://opendope.org/conditions"
  xmlns:odq="http://opendope.org/questions"
  xmlns:odi="http://opendope.org/components"
  xmlns:odgm="http://opendope.org/SmartArt/DataHierarchy"
  xmlns:ns24="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
  xmlns:ns25="http://schemas.openxmlformats.org/drawingml/2006/compatibility"
  xmlns:ns26="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas">
  <w:body>
    <w:altChunk r:id="htmlChunk" />
    <w:sectPr>
      <w:pgSz w:w="<%= width %>" w:h="<%= height %>" w:orient="<%= orient %>" />
      <w:pgMar w:top="<%= margins.top %>"
               w:right="<%= margins.right %>"
               w:bottom="<%= margins.bottom %>"
               w:left="<%= margins.left %>"
               w:header="<%= margins.header %>"
               w:footer="<%= margins.footer %>"
               w:gutter="<%= margins.gutter %>"/>
    </w:sectPr>
  </w:body>
</w:document>
`);

utils = require('./utils');

_ = {
  merge: require('lodash.merge')
};

module.exports = {
  generateDocument: function(zip) {
    var buffer;
    buffer = zip.generate({
      type: 'arraybuffer'
    });
    if (global.Blob) {
      return new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
    } else if (global.Buffer) {
      return new Buffer(new Uint8Array(buffer));
    } else {
      throw new Error("Neither Blob nor Buffer are accessible in this environment. " + "Consider adding Blob.js shim");
    }
  },
  renderDocumentFile: function(documentOptions) {
    var templateData;
    if (documentOptions == null) {
      documentOptions = {};
    }
    templateData = _.merge({
      margins: {
        top: 1440,
        right: 1440,
        bottom: 1440,
        left: 1440,
        header: 720,
        footer: 720,
        gutter: 0
      }
    }, (function() {
      switch (documentOptions.orientation) {
        case 'landscape':
          return {
            height: 12240,
            width: 15840,
            orient: 'landscape'
          };
        default:
          return {
            width: 12240,
            height: 15840,
            orient: 'portrait'
          };
      }
    })(), {
      margins: documentOptions.margins
    });
    return documentTemplate(templateData);
  },
  addFiles: function(zip, htmlSource, documentOptions) {
    zip.file('[Content_Types].xml', fs.readFileSync(__dirname + '/assets/content_types.xml'));
    zip.folder('_rels').file('.rels', fs.readFileSync(__dirname + '/assets/rels.xml'));
    return zip.folder('word').file('document.xml', this.renderDocumentFile(documentOptions)).file('afchunk.mht', utils.getMHTdocument(htmlSource)).folder('_rels').file('document.xml.rels', fs.readFileSync(__dirname + '/assets/document.xml.rels'));
  }
};
