let JSZip = require('jszip');
let fs = require('fs');
var _ = require('underscore');

let html_docx = {asBlob, getMHTdocument, _prepareImageParts, generateDocument, addFiles, renderDocumentFile};

let documentTemplate = _.template(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

let mhtDocumentTemplate = _.template(`MIME-Version: 1.0
Content-Type: multipart/related;
    type="text/html";
    boundary="----=mhtDocumentPart"


------=mhtDocumentPart
Content-Type: text/html;
    charset="utf-8"
Content-Transfer-Encoding: quoted-printable
Content-Location: file:///C:/fake/document.html

<%= htmlSource %>

<%= contentParts %>

------=mhtDocumentPart--
`);

let mhtPartTemplate = _.template(`------=mhtDocumentPart
Content-Type: <%= contentType %>
Content-Transfer-Encoding: <%= contentEncoding %>
Content-Location: <%= contentLocation %>

<%= encodedContent %>
`);

_ = {merge: require('lodash.merge')};

function asBlob(html, options) {
  let zip = new JSZip();
  addFiles(zip, html, options);
  return generateDocument(zip);
}

function generateDocument(zip) {
  var buffer = zip.generate({
    type: 'arraybuffer'
  });
  
  if (global.Blob) {
    return new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
  }
  else if (global.Buffer) {
    return new Buffer(new Uint8Array(buffer));
  } else {
    throw new Error("Neither Blob nor Buffer are accessible in this environment. " + "Consider adding Blob.js shim");
  }
}

function addFiles(zip, htmlSource, documentOptions) {
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType=
        "application/vnd.openxmlformats-package.relationships+xml" />
      <Override PartName="/word/document.xml" ContentType=
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/afchunk.mht" ContentType="message/rfc822"/>
    </Types>
    `);
    
    zip.folder('_rels').file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
          Target="/word/document.xml" Id="R09c83fafc067488e" />
    </Relationships>
    `);
    
    return zip.folder('word')
      .file('document.xml', html_docx.renderDocumentFile(documentOptions))
      .file('afchunk.mht', getMHTdocument(htmlSource))
      .folder('_rels')
      .file('document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk"
        Target="/word/afchunk.mht" Id="htmlChunk" />
    </Relationships>
    `);
  }

function renderDocumentFile(documentOptions) {
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
}

function getMHTdocument(htmlSource) {
    // take care of images
    var imageContentParts, ref;
    ref = _prepareImageParts(htmlSource), htmlSource = ref.htmlSource, imageContentParts = ref.imageContentParts;
    // for proper MHT parsing all '=' signs in html need to be replaced with '=3D'
    htmlSource = htmlSource.replace(/\=/g, '=3D');
    return mhtDocumentTemplate({
        htmlSource: htmlSource,
        contentParts: imageContentParts.join('\n')
    });
}

function _prepareImageParts(htmlSource) {
    var imageContentParts, inlinedReplacer, inlinedSrcPattern;
    imageContentParts = [];
    inlinedSrcPattern = /"data:(\w+\/\w+);(\w+),(\S+)"/g;
    // replacer function for images sources via DATA URI
    inlinedReplacer = function(match, contentType, contentEncoding, encodedContent) {
        var contentLocation, extension, index;
        index = imageContentParts.length;
        extension = contentType.split('/')[1];
        contentLocation = "file:///C:/fake/image" + index + "." + extension;
        imageContentParts.push(mhtPartTemplate({
        contentType: contentType,
        contentEncoding: contentEncoding,
        contentLocation: contentLocation,
        encodedContent: encodedContent
        }));
        return "\"" + contentLocation + "\"";
    };
    if (typeof htmlSource === 'string') {
        if (!/<img/g.test(htmlSource)) {
        return {
            htmlSource: htmlSource,
            imageContentParts: imageContentParts
        };
        }
        htmlSource = htmlSource.replace(inlinedSrcPattern, inlinedReplacer);
        return {
          htmlSource: htmlSource,
          imageContentParts: imageContentParts
        };
    } else {
        throw new Error("Not a valid source provided!");
    }
}

module.exports = html_docx;
