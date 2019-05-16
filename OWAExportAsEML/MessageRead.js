(function () {
  "use strict";

  var messageBanner;

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function () {
    $(document).ready(function () {
         var request = GetItem();
        var envelope = getSoapEnvelope(request);
        Office.context.mailbox.makeEwsRequestAsync(envelope, function(result){
         var parser = new DOMParser();
         var doc = parser.parseFromString(result.value, "text/xml");
         var values = doc.getElementsByTagName("t:MimeContent");
         var subject = doc.getElementsByTagName("t:Subject");
         console.log(subject[0].textContent)
         download((subject[0].textContent + ".eml"),values[0].textContent);
        });
    });
};

function base64toBlob(base64Data, contentType) {
  contentType = contentType || '';
  var sliceSize = 1024;
  var byteCharacters = atob(base64Data);
  var bytesLength = byteCharacters.length;
  var slicesCount = Math.ceil(bytesLength / sliceSize);
  var byteArrays = new Array(slicesCount);

  for (var sliceIndex = 0; sliceIndex < slicesCount; ++sliceIndex) {
      var begin = sliceIndex * sliceSize;
      var end = Math.min(begin + sliceSize, bytesLength);

      var bytes = new Array(end - begin);
      for (var offset = begin, i = 0; offset < end; ++i, ++offset) {
          bytes[i] = byteCharacters[offset].charCodeAt(0);
      }
      byteArrays[sliceIndex] = new Uint8Array(bytes);
  }
  return new Blob(byteArrays, { type: contentType });
}

function download(filename, text) {
  var downloadblob = base64toBlob(text);
  if (window.navigator && window.navigator.msSaveOrOpenBlob) {
    window.navigator.msSaveOrOpenBlob(downloadblob,filename);
    return;
  } 
  const data = window.URL.createObjectURL(downloadblob);
  var element = document.createElement('a');
  element.setAttribute('href', data);
  element.setAttribute('download', filename);
  element.style.display = 'none';
  document.body.appendChild(element);
  element.click();
  document.body.removeChild(element);
}
function getSoapEnvelope(request) {
    // Wrap an Exchange Web Services request in a SOAP envelope.
    var result =

    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +

    request +

    '  </soap:Body>' +
    '</soap:Envelope>';

    return result;
}

function GetItem() {
    var results =
  '  <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
  '    <ItemShape>' +
  '      <t:BaseShape>IdOnly</t:BaseShape>' +
  '      <t:IncludeMimeContent>true</t:IncludeMimeContent>' +
  '      <AdditionalProperties xmlns="http://schemas.microsoft.com/exchange/services/2006/types">' +
  '        <FieldURI FieldURI="item:Subject" />' +
  '      </AdditionalProperties>' +
  '    </ItemShape>' +
  '    <ItemIds>' +
  '      <t:ItemId Id="' + Office.context.mailbox.item.itemId + '" />' +
  '    </ItemIds>' +
  '  </GetItem>';
 
    return results;
}



 
  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();