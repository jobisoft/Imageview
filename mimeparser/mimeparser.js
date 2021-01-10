/**
 * https://theofficecontext.com/2019/05/24/simple-client-side-javascript-mime-parser/
 * Simple MIME Object for parsing basic MIME results:
 * - Headers (parts)
 * - HTML (full)
 * - Attachments
 * - Embedded Image Data
 *
 * Modifications by John Bieling
 * - include filename
 * - include ContentType
 */
function simpleMIMEClass() {
  /** @type {string} */
  var fullHeader = "";
  /** @type {string} */
  var html = "";
  /** @type {MIMEImageData[]} */
  var images = new Array();
  /** @type {string} */
  var text = "";
  /** @type {string} */
  var mimeData = "";

  /** 
   * Internal MIME Image Class
  */
  function MIMEImageData(){
    this.ID = "";
    this.filename = "";
    this.imageData = "";
    this.contentType = "";
  }

  /** @type {string} */
  this.FullHeader = function() { return fullHeader };
  
  /** @type {string} */
  this.HTML = function() { return html; }

  /** @type {string} */
  this.Text = function() { return text; }

  /** @type {string} */
  this.GetImageData = function() {
    return images;
  }

  /**
   * Parses the MIME data into the object
   * @param {@type {string} value} – MIME data from getMailItemMimeContent()
   */
  this.Parse = function(value) {
    mimeData = value;
    text = value;//atob(mimeData);
    var parts = text.split("\r\n\r\n");
    fullHeader = parts[0];
    for(var partIndex=0;partIndex<parts.length;partIndex++) {
      if(parts[partIndex].includes("Content-Type: text/html;", 0) > 0) {
        html = parts[partIndex+1];
        // must remove the =3D which is an incomplete escape code for "="
        // which gets into Outlook MIME somehow – now sure
        // also removing line breaks which are = at the very end
        // followed by carriage return and new line
        html = html.replace(/=3D/g,'=').replace(/=\r\n/g,"");
      }
      if(parts[partIndex].includes("Content-Type: image/", 0) > 0) {
        var imgTag = parts[partIndex].split("\r\n");
        var imgData = parts[partIndex+1].split("\r\n--")[0];
        
        let contentTypeArray = imgTag.filter(e => e.includes("Content-Type: "));
        let contentType = (contentTypeArray.length > 0) 
          ? contentTypeArray[0].split("Content-Type: ").pop()
          : "";
        while (contentType.endsWith(";")) {
          contentType = contentType.slice(0, -1);
        }
        
        let filenameArray = imgTag.filter(e => e.includes("filename="));
        let filename = (filenameArray.length > 0) 
          ? filenameArray[0].split("filename=").pop().replaceAll('\"','')
          : "";
        while (filename.endsWith(";")) {
          filename = filename.slice(0, -1);
        }

        var imgID = "";
        for(var tagIndex=0;tagIndex<imgTag.length;tagIndex++) {
          if(imgTag[tagIndex].includes("Content-ID: ") > 0) {
            imgID = "cid:" + imgTag[tagIndex].split(": ")[1].replace("<","").replace(">","");
          }
        }
        var img = new MIMEImageData();
        img.ID = imgID;
        img.imageData = imgData;
        img.filename = filename;
        img.contentType = contentType;
        images.push(img);
      }
    }

    // done
    return this;
  };
};