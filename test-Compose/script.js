// This function is called when Office.js is ready to start your Add-in
var _mailbox;
var _Item;
var _AppGuid = "99429ef8-be83-4ce2-ba79-f4471f89f674";
var _ItemGuid = "";

Office.initialize = function () {
    $(document).ready(function () {
     
        _ItemGuid = guid();
        var item = Office.context.mailbox.item;
        _Item = item;
        _Item.loadCustomPropertiesAsync(customPropsCallback);
    

    });
};
function saveCallback(asyncResult) {
    _Item.saveAsync(saveItemCallBack);
   ;
}

function saveItemCallBack(asyncResult) {
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
    $('#ChkTest').text(asyncResult.value);
    var result = asyncResult.value;
    var context = asyncResult.context;
    var is_chrome = navigator.userAgent.toLowerCase().indexOf('chrome') > -1;
    if (is_chrome) {
        var parser = new DOMParser();
        var doc = parser.parseFromString(asyncResult.value, "text/xml");
        var values = doc.childNodes[0].getElementsByTagName("ItemId");
        $('#ChkTest').text(values[0].attributes['Id'].value);
        $('#ChkTest').text(Base64.encode(getVerbStream()));
        
    }
    else {
        var parser = new DOMParser();
        var doc = parser.parseFromString(asyncResult.value, "text/xml");
        var values = doc.childNodes[0].getElementsByTagName("t:ItemId");
        $('#ChkTest').text(values[0].attributes['Id'].value);
        $('#ChkTest').text(Base64.encode(getVerbStream()));
    }
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

function getVerbStream() {
    var Header = "02010600000000000000";
    var ReplyToAllHeader = "055265706C790849504D2E4E6F7465074D657373616765025245050000000000000000";
    var ReplyToAllFooter = "0000000000000002000000660000000200000001000000";
    var ReplyToHeader = "0C5265706C7920746F20416C6C0849504D2E4E6F7465074D657373616765025245050000000000000000";
    var ReplyToFooter = "0000000000000002000000670000000300000002000000";
    var ForwardHeader = "07466F72776172640849504D2E4E6F7465074D657373616765024657050000000000000000";
    var ForwardFooter = "0000000000000002000000680000000400000003000000";
    var ReplyToFolderHeader = "0F5265706C7920746F20466F6C6465720849504D2E506F737404506F737400050000000000000000";
    var ReplyToFolderFooter = "00000000000000020000006C00000008000000";
    var ApproveOption = "0400000007417070726F76650849504D2E4E6F74650007417070726F766500000000000000000001000000020000000200000001000000FFFFFFFF";
    var RejectOption= "040000000652656A6563740849504D2E4E6F7465000652656A65637400000000000000000001000000020000000200000002000000FFFFFFFF";
    var VoteOptionExtras = "0401055200650070006C00790002520045000C5200650070006C007900200074006F00200041006C006C0002520045000746006F007200770061007200640002460057000F5200650070006C007900200074006F00200046006F006C00640065007200000741007000700072006F00760065000741007000700072006F007600650006520065006A0065006300740006520065006A00650063007400";
    var DisableReplyAllVal = "00";
    var DisableReplyAllVal = "01";
    var DisableReplyVal = "00";
    var DisableReplyVal = "01";
    var DisableForwardVal = "00";
    var DisableForwardVal = "01";
    var DisableReplyToFolderVal = "00";
    var DisableReplyToFolderVal = "01";
    var VerbValue = Header + ReplyToAllHeader + DisableReplyAllVal + ReplyToAllFooter + ReplyToHeader + DisableReplyVal + ReplyToFooter + ForwardHeader + DisableForwardVal + ForwardFooter + ReplyToFolderHeader + DisableReplyToFolderVal + ReplyToFolderFooter + ApproveOption  + RejectOption + VoteOptionExtras;
    return VerbValue;
}

function UpdateVerb(Id, ChangeKey, Value) {
    var results =

   ' <UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite" SendMeetingInvitationsOrCancellations="SendToNone" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
   '         <ItemChanges>' +
   '           <t:ItemChange>' +
   '            <t:ItemId Id="' + Id + '" ChangeKey="' + ChangeKey + '" />' +
   '             <t:Updates>' +
   '               <t:SetItemField>' +
   '                 <t:ExtendedFieldURI DistinguishedPropertySetId="Common" PropertyId="34080" PropertyType="Integer" />' +
   '                 <t:Message>' +
   '                   <t:ExtendedProperty>' +
   '                   <t:ExtendedFieldURI DistinguishedPropertySetId="Common" PropertyId="34080" PropertyType="Integer" />' +
   '                   <t:Value>' + Value + '</t:Value>' +
   '                  </t:ExtendedProperty>' +
   '                 </t:Message>' +
   '               </t:SetItemField>' +
   '             </t:Updates>' +
   '           </t:ItemChange>' +
   '         </ItemChanges>' +
   '</UpdateItem>';
    return results;
}

// Create Base64 Object
var Base64 = { _keyStr: "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=", encode: function (e) { var t = ""; var n, r, i, s, o, u, a; var f = 0; e = Base64._utf8_encode(e); while (f < e.length) { n = e.charCodeAt(f++); r = e.charCodeAt(f++); i = e.charCodeAt(f++); s = n >> 2; o = (n & 3) << 4 | r >> 4; u = (r & 15) << 2 | i >> 6; a = i & 63; if (isNaN(r)) { u = a = 64 } else if (isNaN(i)) { a = 64 } t = t + this._keyStr.charAt(s) + this._keyStr.charAt(o) + this._keyStr.charAt(u) + this._keyStr.charAt(a) } return t }, decode: function (e) { var t = ""; var n, r, i; var s, o, u, a; var f = 0; e = e.replace(/[^A-Za-z0-9\+\/\=]/g, ""); while (f < e.length) { s = this._keyStr.indexOf(e.charAt(f++)); o = this._keyStr.indexOf(e.charAt(f++)); u = this._keyStr.indexOf(e.charAt(f++)); a = this._keyStr.indexOf(e.charAt(f++)); n = s << 2 | o >> 4; r = (o & 15) << 4 | u >> 2; i = (u & 3) << 6 | a; t = t + String.fromCharCode(n); if (u != 64) { t = t + String.fromCharCode(r) } if (a != 64) { t = t + String.fromCharCode(i) } } t = Base64._utf8_decode(t); return t }, _utf8_encode: function (e) { e = e.replace(/\r\n/g, "\n"); var t = ""; for (var n = 0; n < e.length; n++) { var r = e.charCodeAt(n); if (r < 128) { t += String.fromCharCode(r) } else if (r > 127 && r < 2048) { t += String.fromCharCode(r >> 6 | 192); t += String.fromCharCode(r & 63 | 128) } else { t += String.fromCharCode(r >> 12 | 224); t += String.fromCharCode(r >> 6 & 63 | 128); t += String.fromCharCode(r & 63 | 128) } } return t }, _utf8_decode: function (e) { var t = ""; var n = 0; var r = c1 = c2 = 0; while (n < e.length) { r = e.charCodeAt(n); if (r < 128) { t += String.fromCharCode(r); n++ } else if (r > 191 && r < 224) { c2 = e.charCodeAt(n + 1); t += String.fromCharCode((r & 31) << 6 | c2 & 63); n += 2 } else { c2 = e.charCodeAt(n + 1); c3 = e.charCodeAt(n + 2); t += String.fromCharCode((r & 15) << 12 | (c2 & 63) << 6 | c3 & 63); n += 3 } } return t } }

