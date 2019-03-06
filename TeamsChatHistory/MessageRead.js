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
                    if(Office.context.mailbox.item.sender.emailAddress == "noreply@email.teams.microsoft.com"){
                        var senderName = Office.context.mailbox.item.sender.displayName.replace(' in Teams','');
                        resolveName(accessToken,senderName); 
                    }else{
                        getChatMessages(accessToken);
                    }                   
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

    function resolveName(accessToken,NameToLookup){
        var request = GetResolveNameRequest();
        var envelope = getSoapEnvelope(request);
        Office.context.mailbox.makeEwsRequestAsync(envelope, function (asyncResult,accessToken) {
            console.log(asyncResult);
            console.log(accessToken);

        });

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
    
    function GetResolveNameRequest(NameToLookup) {
        var results =    
        '<ResolveNames xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types "ReturnFullContactData="true">' +
        '        <UnresolvedEntry>' + NameToLookup + '</UnresolvedEntry>' +
        ' </ResolveNames>';
         return results;
    }
    
    function DisplayMessages(Messages) {
        try {
            var html = "<div class=\"ms-Table-row\">";
            html = html + "<span class=\"ms-Table-cell\" >ReceivedDateTime</span>";
            html = html + "<span class=\"ms-Table-cell\">BodyPreview</span>";
            html = html + "</div>";
            var i;
            for (i = (Messages.length-1); i >= 0 ; i--) { 
                var rcvDate = Date.parse(Messages[i].ReceivedDateTime);
                html = html + "<div class=\"ms-Table-row\">";
                html = html +"<span class=\"ms-Table-cell ms-fontWeight-semibold\">" + rcvDate.toString('dd-MMM HH:mm') + "</span>";
                html = html +"<span id=\"Subject\" class=\"ms-Table-cell\">";
                html = html + Messages[i].BodyPreview + " <a target='_blank' href='" + Messages[i].WebLink + "'> Link</a></span ></div >";
            }

            //Messages.forEach(function (Message) {

            //});
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