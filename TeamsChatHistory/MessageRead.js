﻿(function () {
    "use strict";

    var messageBanner;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
                if (result.status === "succeeded") {
                    var accessToken = result.value;
                    getChatMessages(accessToken);
                } else {
                    // Handle the error
                }
            });

        });

    };

    function getChatMessages(accessToken) {
        var filterString = "SingleValueExtendedProperties/Any(ep: ep/PropertyId eq 'String 0x001a' and ep/Value eq 'IPM.SkypeTeams.Message') and from/emailAddress/address eq '" + Office.context.mailbox.item.sender.emailAddress + "'";
        var GetURL = "https://outlook.office.com/api/v2.0/me/MailFolders/AllItems/messages?$Top=100&$Select=ReceivedDateTime,bodyPreview,webLink&$filter=" + filterString;
        $.ajax({
            type: "Get",
            contentType: "application/json; charset=utf-8",
            url: GetURL,
            dataType: 'json',
            headers: { 'Authorization': 'Bearer ' + accessToken }
        }).done(function (item) {
            DisplayMessages(item.value);
        }).fail(function (error) {
            $('#mTchatTable').append("Error getting Messages " + error);
        });
    }

    function DisplayMessages(Messages) {
        try {
            var html = "<div class=\"ms-Table-row\">";
            html = html + "<span class=\"ms-Table-cell\" >ReceivedDateTime</span>";
            html = html + "<span class=\"ms-Table-cell\">BodyPreview</span>";
            html = html + "</div>";
            var msgTable = ""
            Messages.forEach(function (Message) {
                var rcvDate = Date(Date.parse(Message.ReceivedDateTime));
                var rcvDateString = rcvDate.toLocaleString('GB', {
                    'day': 'numeric',
                    'hour': '2-digit',
                    'minute': '2-digit',
                    'hour12': false,
                    'month': 'long'
                });
                var tablerow = "<div class=\"ms-Table-row\">";
                tablerow =  tablerow + "<span class=\"ms-Table-cell ms-fontWeight-semibold\">" + rcvDateString + "</span>";
                tablerow =  tablerow + "<span id=\"Subject\" class=\"ms-Table-cell\">"+ html;
                tablerow =  tablerow + Message.BodyPreview + " <a target='_blank' href='" + Message.WebLink + "'> Link</a></span ></div >";
                msgTable = tablerow + msgTable;
            });
            $('#mTchatTable').append((html + msgTable));
        } catch (error) {
            $('#mTchatTable').html("Error displaying table " + error);
        }
    }


    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();