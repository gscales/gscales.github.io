(function () {
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
            Messages.forEach(function (Message) {
                console.log(Message);
                var rcvDate = Date(Date.parse(Message.ReceivedDateTime));
                html = html + "<div class=\"ms-Table-row\">";
                html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\">" + Message.ReceivedDateTime + "</span>";
                html = html + "<span id=\"Subject\" class=\"ms-Table-cell\">";
                html = html + Message.BodyPreview + " <a target='_blank' href='" + Message.WebLink + "'> Link</a></span ></div >"
            });
            $('#mTchatTable').append(html);
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