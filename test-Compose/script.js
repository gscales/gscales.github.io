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

(function (global) { "use strict"; var _Base64 = global.Base64; var version = "2.1.9"; var buffer; if (typeof module !== "undefined" && module.exports) { try { buffer = require("buffer").Buffer } catch (err) { } } var b64chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"; var b64tab = function (bin) { var t = {}; for (var i = 0, l = bin.length; i < l; i++) t[bin.charAt(i)] = i; return t }(b64chars); var fromCharCode = String.fromCharCode; var cb_utob = function (c) { if (c.length < 2) { var cc = c.charCodeAt(0); return cc < 128 ? c : cc < 2048 ? fromCharCode(192 | cc >>> 6) + fromCharCode(128 | cc & 63) : fromCharCode(224 | cc >>> 12 & 15) + fromCharCode(128 | cc >>> 6 & 63) + fromCharCode(128 | cc & 63) } else { var cc = 65536 + (c.charCodeAt(0) - 55296) * 1024 + (c.charCodeAt(1) - 56320); return fromCharCode(240 | cc >>> 18 & 7) + fromCharCode(128 | cc >>> 12 & 63) + fromCharCode(128 | cc >>> 6 & 63) + fromCharCode(128 | cc & 63) } }; var re_utob = /[\uD800-\uDBFF][\uDC00-\uDFFFF]|[^\x00-\x7F]/g; var utob = function (u) { return u.replace(re_utob, cb_utob) }; var cb_encode = function (ccc) { var padlen = [0, 2, 1][ccc.length % 3], ord = ccc.charCodeAt(0) << 16 | (ccc.length > 1 ? ccc.charCodeAt(1) : 0) << 8 | (ccc.length > 2 ? ccc.charCodeAt(2) : 0), chars = [b64chars.charAt(ord >>> 18), b64chars.charAt(ord >>> 12 & 63), padlen >= 2 ? "=" : b64chars.charAt(ord >>> 6 & 63), padlen >= 1 ? "=" : b64chars.charAt(ord & 63)]; return chars.join("") }; var btoa = global.btoa ? function (b) { return global.btoa(b) } : function (b) { return b.replace(/[\s\S]{1,3}/g, cb_encode) }; var _encode = buffer ? function (u) { return (u.constructor === buffer.constructor ? u : new buffer(u)).toString("base64") } : function (u) { return btoa(utob(u)) }; var encode = function (u, urisafe) { return !urisafe ? _encode(String(u)) : _encode(String(u)).replace(/[+\/]/g, function (m0) { return m0 == "+" ? "-" : "_" }).replace(/=/g, "") }; var encodeURI = function (u) { return encode(u, true) }; var re_btou = new RegExp(["[À-ß][-¿]", "[à-ï][-¿]{2}", "[ð-÷][-¿]{3}"].join("|"), "g"); var cb_btou = function (cccc) { switch (cccc.length) { case 4: var cp = (7 & cccc.charCodeAt(0)) << 18 | (63 & cccc.charCodeAt(1)) << 12 | (63 & cccc.charCodeAt(2)) << 6 | 63 & cccc.charCodeAt(3), offset = cp - 65536; return fromCharCode((offset >>> 10) + 55296) + fromCharCode((offset & 1023) + 56320); case 3: return fromCharCode((15 & cccc.charCodeAt(0)) << 12 | (63 & cccc.charCodeAt(1)) << 6 | 63 & cccc.charCodeAt(2)); default: return fromCharCode((31 & cccc.charCodeAt(0)) << 6 | 63 & cccc.charCodeAt(1)) } }; var btou = function (b) { return b.replace(re_btou, cb_btou) }; var cb_decode = function (cccc) { var len = cccc.length, padlen = len % 4, n = (len > 0 ? b64tab[cccc.charAt(0)] << 18 : 0) | (len > 1 ? b64tab[cccc.charAt(1)] << 12 : 0) | (len > 2 ? b64tab[cccc.charAt(2)] << 6 : 0) | (len > 3 ? b64tab[cccc.charAt(3)] : 0), chars = [fromCharCode(n >>> 16), fromCharCode(n >>> 8 & 255), fromCharCode(n & 255)]; chars.length -= [0, 0, 2, 1][padlen]; return chars.join("") }; var atob = global.atob ? function (a) { return global.atob(a) } : function (a) { return a.replace(/[\s\S]{1,4}/g, cb_decode) }; var _decode = buffer ? function (a) { return (a.constructor === buffer.constructor ? a : new buffer(a, "base64")).toString() } : function (a) { return btou(atob(a)) }; var decode = function (a) { return _decode(String(a).replace(/[-_]/g, function (m0) { return m0 == "-" ? "+" : "/" }).replace(/[^A-Za-z0-9\+\/]/g, "")) }; var noConflict = function () { var Base64 = global.Base64; global.Base64 = _Base64; return Base64 }; global.Base64 = { VERSION: version, atob: atob, btoa: btoa, fromBase64: decode, toBase64: encode, utob: utob, encode: encode, encodeURI: encodeURI, btou: btou, decode: decode, noConflict: noConflict }; if (typeof Object.defineProperty === "function") { var noEnum = function (v) { return { value: v, enumerable: false, writable: true, configurable: true } }; global.Base64.extendString = function () { Object.defineProperty(String.prototype, "fromBase64", noEnum(function () { return decode(this) })); Object.defineProperty(String.prototype, "toBase64", noEnum(function (urisafe) { return encode(this, urisafe) })); Object.defineProperty(String.prototype, "toBase64URI", noEnum(function () { return encode(this, true) })) } } if (global["Meteor"]) { Base64 = global.Base64 } })(this);
