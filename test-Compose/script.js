// This function is called when Office.js is ready to start your Add-in
var _mailbox;
var _Item;

Office.initialize = function () {
    $(document).ready(function () {
    var item = Office.context.mailbox.item;
    $('#ChkTest').text(item.itemId);
        //var request = getItemRequest(_Item.itemId);
        //var envelope = getSoapEnvelope(request);
        
        //_mailbox.makeEwsRequestAsync(envelope, callbackGetItem);
    });
};


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