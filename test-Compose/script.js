// This function is called when Office.js is ready to start your Add-in
var _mailbox;
var _Item;
var _AppGuid = "99429ef8-be83-4ce2-ba79-f4471f89f674";

Office.initialize = function () {
    $(document).ready(function () {
        var item = Office.context.mailbox.item;
        var request = FindItemRequest();
        var envelope = getSoapEnvelope(request);
        _mailbox.makeEwsRequestAsync(envelope, callbackFindItems);
        //var request = getItemRequest(_Item.itemId);
        //var envelope = getSoapEnvelope(request);
        
        //_mailbox.makeEwsRequestAsync(envelope, callbackGetItem);
    });
};
function saveCallback(asyncResult) {
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
 '           <t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName=cecp-"' + _AppGuid + ' PropertyType="String" />' +
 '        </t:AdditionalProperties>' +
 '       </m:ItemShape>' +
 '       <m:IndexedPageItemView MaxEntriesReturned="100" Offset="0" BasePoint="Beginning" />' +
 '       <m:Restriction>' +
 '         <t:Exists>' +
 '           <t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName=cecp-"' + _AppGuid + ' PropertyType="String" />' +
 '         </t:Exists>' +
 '       </m:Restriction>' +
 '       <m:ParentFolderIds>' +
 '         <t:DistinguishedFolderId Id="drafts" />' +
 '       </m:ParentFolderIds>' +
 '     </m:FindItem>';
    return result;
}

function addEmoticon(Emoticon) {
    if ($('#BodyRadio').is(':checked')) {
        AddEmoticonToBody(Emoticon);
    }
    else {
        AddEmoticonToSubject(Emoticon);
    }
    
}
function AddEmoticonToSubject(Emoticon) {
    var item = Office.context.mailbox.item;
    item.subject.getAsync(
    function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            //write(asyncResult.error.message);
        }
        else {
            item.subject.setAsync(asyncResult.value + Emoticon);
        }
    });

}
function AddEmoticonToBody(Emoticon) {
    var item = Office.context.mailbox.item;
    item.body.getTypeAsync(
         function (result) {
             if (result.status == Office.AsyncResultStatus.Failed){
                 write(result.error.message);
             }
             else {
                 // Successfully got the type of item body.
                 // Set data of the appropriate type in body.
                 if (result.value == Office.MailboxEnums.BodyType.Html) {
                     // Body is of HTML type.
                     // Specify HTML in the coercionType parameter
                     // of setSelectedDataAsync.
                     //********************* 
                     //Note CoercionType has been set to Text as a workaround
                     //********************
                     item.body.setSelectedDataAsync(
                        Emoticon,
                         { coercionType: Office.CoercionType.Text, 
                             asyncContext: { var3: 1, var4: 2 } },
                         function (asyncResult) {
                             if (asyncResult.status == 
                                 Office.AsyncResultStatus.Failed){
                                 write(asyncResult.error.message);
                             }
                             else {
                                 // Successfully set data in item body.
                                 // Do whatever appropriate for your scenario,
                                 // using the arguments var3 and var4 as applicable.
                             }
                         });
                 }
                 else {
                     // Body is of text type. 
                     item.body.setSelectedDataAsync(
                         Emoticon,
                         { coercionType: Office.CoercionType.Text, 
                             asyncContext: { var3: 1, var4: 2 } },
                         function (asyncResult) {
                             if (asyncResult.status == 
                                 Office.AsyncResultStatus.Failed){
                                 write(asyncResult.error.message);
                             }
                             else {
                                 // Successfully set data in item body.
                                 // Do whatever appropriate for your scenario,
                                 // using the arguments var3 and var4 as applicable.
                             }
                         });
                 }
             }
         });
}
function BuildEmoticonTable() {
    var Emoticons = [
    "╚═། ◑ ▃ ◑ །═╝",
    "¯\_(ツ)_/¯",
    "  o͡͡͡╮░ O ◡ O ░╭o͡͡͡ ",
    "ʘ ͜ʖ ʘ",
    "ᕙ(▀̿̿Ĺ̯̿̿▀̿ ̿) ᕗ",
    "ᕕ(⌐■_■)ᕗ ♪♬",
    "║ ಡ ͜ ʖ ಡ ║",
    "ᕕ( ՞ ᗜ ՞ )ᕗ",
    "ლ(ಠ益ಠ)ლ",
    "(ಠ_ಠ)",
    "(╯_╰)",
    "(ﾉﾟ0ﾟ)ﾉ",
    "( •_•)O*¯`·.¸.·´¯`°Q(•_• )",
    " ♪♫•*¨*•.¸¸❤¸¸.•*¨*•♫♪ ",
    " •*´¨`*•.¸¸.•*´¨`*•.¸¸. ",
    " (ᵔᴥᵔ) ",
    " 눈_눈 ",
    " \(*0*)/ ",
    " {♥‿ ♥} ",
    " [̲̅$̲̅(̲̅ιοο̲̅)̲̅$̲̅] "
    ];

    var $table = $('<table cellspacing="3" class="IconTable" />');
  
    for (index = 0; index < Emoticons.length; index++) {
        var $NewRow = $('<tr />').appendTo($table);
        $('<td />').html('<a class="auto-style2" href="#" onclick="addEmoticon(\'' + Emoticons[index] + '\'); return false;">' + Emoticons[index] + '</a>').appendTo($NewRow);
        index++;
        if (index < Emoticons.length) {
            $('<td />').html('<a class="auto-style2" href="#" onclick="addEmoticon(\'' + Emoticons[index] + '\'); return false;">' + Emoticons[index] + '</a>').appendTo($NewRow);
        }
    }
    $table.appendTo($('#Icons'));
}