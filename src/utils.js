var mhtDocumentTemplate, mhtPartTemplate;
var _ = require('underscore');

mhtDocumentTemplate = _.template(`MIME-Version: 1.0
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

mhtPartTemplate = _.template(`------=mhtDocumentPart
Content-Type: <%= contentType %>
Content-Transfer-Encoding: <%= contentEncoding %>
Content-Location: <%= contentLocation %>

<%= encodedContent %>
`);

module.exports = {
  getMHTdocument: function(htmlSource) {
    // take care of images
    var imageContentParts, ref;
    ref = this._prepareImageParts(htmlSource), htmlSource = ref.htmlSource, imageContentParts = ref.imageContentParts;
    // for proper MHT parsing all '=' signs in html need to be replaced with '=3D'
    htmlSource = htmlSource.replace(/\=/g, '=3D');
    return mhtDocumentTemplate({
      htmlSource: htmlSource,
      contentParts: imageContentParts.join('\n')
    });
  },
  _prepareImageParts: function(htmlSource) {
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
};
