// This function is called when Office.js is ready to start your Add-in
Office.initialize = function (reason) { 
	$(document).ready(function () {
	    BuildEmoticonTable();
	});
}; 

function addEmoticon(Emoticon) {
    AddEmoticonToBody(Emoticon);
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
    item.body.getAsync(
    function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            //write(asyncResult.error.message);
        }
        else {
            item.body.setAsync(asyncResult.value + Emoticon);
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
    "♪♫•*¨*•.¸¸❤¸¸.•*¨*•♫♪",
    "•*´¨`*•.¸¸.•*´¨`*•.¸¸.",
    "(ᵔᴥᵔ)",
    "눈_눈",
    "\(*0*)/",
    "{♥‿ ♥}",
    "[̲̅$̲̅(̲̅ιοο̲̅)̲̅$̲̅]"
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