// This function is called when Office.js is ready to start your Add-in
var _mailbox;
var _Item;
var _AppGuid = "99429ef8-be83-4ce2-ba79-f4471f89f674";
var _ItemGuid = "";

Office.initialize = function () {
    $(document).ready(function () {
        _ItemGuid = guid();
        var item = Office.context.mailbox.item;
        item.loadCustomPropertiesAsync(customPropsCallback);
        

        //var request = getItemRequest(_Item.itemId);
        //var envelope = getSoapEnvelope(request);
        
        //_mailbox.makeEwsRequestAsync(envelope, callbackGetItem);
    });
};
function saveCallback(asyncResult) {
    var request = FindItemRequest();
    var envelope = getSoapEnvelope(request);
    $('#ChkTest').text(request);
    Office.context.mailbox.makeEwsRequestAsync(envelope, callbackFindItems);
}

function guid() {
    function s4() {
        return Math.floor((1 + Math.random()) * 0x10000)
          .toString(16)
          .substring(1);
    }
    return s4() + s4() + '-' + s4() + '-' + s4() + '-' +
      s4() + '-' + s4() + s4() + s4();
}

function callbackFindItems(asyncResult) {
    var result = asyncResult.value;
    var context = asyncResult.context;
    $('#ChkTest').text(result);
}
function getSoapEnvelope(request) {
    // Wrap an Exchange Web Services request in a SOAP envelope.
    var result =

    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
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
function FindItemRequest() {
    // Return a GetItem operation request for the subject of the specified item. 
    var result =
 '   <m:FindItem Traversal="Shallow">' +
 '     <m:ItemShape>' +
 '     <t:BaseShape>IdOnly</t:BaseShape>' +
 '        <t:AdditionalProperties>' +
 '           <t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="cecp-' + _AppGuid + '" PropertyType="String" />' +
 '        </t:AdditionalProperties>' +
 '       </m:ItemShape>' +
 '       <m:IndexedPageItemView MaxEntriesReturned="100" Offset="0" BasePoint="Beginning" />' +
 '       <m:Restriction>' +
 '          <t:IsEqualTo>' +
 '              <t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="cecp-' + _AppGuid + '" PropertyType="String" />' +
 '              <t:FieldURIOrConstant>' +
 '                 <t:Constant Value="{&quot;nssplugIn&quot;:&quot;' + _ItemGuid + '&quot;}" />' +
 '              </t:FieldURIOrConstant>' +
 '           </t:IsEqualTo>' +
 '       </m:Restriction>' +
 '       <m:ParentFolderIds>' +
 '         <t:DistinguishedFolderId Id="drafts" />' +
 '       </m:ParentFolderIds>' +
 '     </m:FindItem>';
    return result;
}

function customPropsCallback(asyncResult) {
    var customProps = asyncResult.value;
    customProps.set("nssplugIn", _ItemGuid);
    customProps.saveAsync(saveCallback);
}