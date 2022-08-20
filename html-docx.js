let JSZip = require('jszip');
let fs = require('fs');
var _ = require('underscore');
let assert = require('assert');

let html_docx = {asBlob, getMHTdocument, _prepareImageParts, generateDocument, addFiles, renderDocumentFile};

async function asBlob(html, options) {
  let zip = new JSZip();
  addFiles(zip, html, options);
  return await generateDocument(zip);
}

function addFiles(zip, html, documentOptions) {
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/afchunk.mht" ContentType="message/rfc822"/>
    </Types>`);
    
  zip.folder('_rels').file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml" Id="R09c83fafc067488e" />
    </Relationships>`);
    
  return zip.folder('word')
    .file('document.xml', html_docx.renderDocumentFile(documentOptions))
    .file('afchunk.mht', getMHTdocument(html))
    .folder('_rels')
    .file('document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="/word/afchunk.mht" Id="htmlChunk" />
      </Relationships>`);
  }

async function generateDocument(zip) {
  assert(global.Buffer, 'Buffer not available, is html-docx not being run in Node.js?');

  var buffer = await zip.generateAsync({
    type: 'arraybuffer'
  });  
  return Buffer.from(new Uint8Array(buffer));
}

function renderDocumentFile(opts) {
    if (!opts)
      opts = {};

    if (opts.orientation == 'landscape') {
      data = {
        height: 12240,
        width: 15840,
        orient: 'landscape'
      };
    } else {
      data = {
        width: 12240,
        height: 15840,
        orient: 'portrait'
      };
    }

    data.margins = Object.assign({
      top: 1440,
      right: 1440,
      bottom: 1440,
      left: 1440,
      header: 720,
      footer: 720,
      gutter: 0
    }, opts.margins || {});
        
    return _.template(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
`)(data);
}

function getMHTdocument(original_html) {
  // take care of images
  let {html, imageContentParts} = _prepareImageParts(original_html);
  
  // for proper MHT parsing all '=' signs in html need to be replaced with '=3D'
  html = html.replace(/\=/g, '=3D');

  return _.template(`MIME-Version: 1.0
Content-Type: multipart/related;
    type="text/html";
    boundary="----=mhtDocumentPart"


------=mhtDocumentPart
Content-Type: text/html;
    charset="utf-8"
Content-Transfer-Encoding: quoted-printable
Content-Location: file:///C:/fake/document.html

<%= html %>

<%= contentParts %>

------=mhtDocumentPart--
`)({html: html, contentParts: imageContentParts.join('\n')});
}

let mhtDocumentPartTemplate = _.template(`------=mhtDocumentPart
Content-Type: <%= contentType %>
Content-Transfer-Encoding: <%= contentEncoding %>
Content-Location: <%= file_url %>

<%= encodedContent %>
`);

function _prepareImageParts(html) {
  if (typeof html != 'string')
    throw new Error('invalid html source passed to html-docx-js _prepareImageParts()');
  if (!/<img/g.test(html))
    return {html: html, imageContentParts: []};

  let imageContentParts = [];
  
  // replace image data: sources
  html = html.replace(/"data:(\w+\/\w+);(\w+),(\S+)"/g, function(match, contentType, contentEncoding, encodedContent) {
    let file_ext = contentType.split('/')[1];
    let file_url = `file:///C:/fake/image${imageContentParts.length}.${file_ext}`;
    imageContentParts.push(mhtDocumentPartTemplate({contentType, contentEncoding, file_url, encodedContent}));
    return `"${file_url}"`;
  });

  return {
    html: html,
    imageContentParts: imageContentParts
  };
}

module.exports = html_docx;
