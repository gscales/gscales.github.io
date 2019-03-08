(function () {
    "use strict";

    var messageBanner;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            if(Office.context.mailbox.item.sender.emailAddress == "noreply@email.teams.microsoft.com"){
                resolveName(Office.context.mailbox.item.sender.displayName.replace(" in Teams",""));
            }else{
                getRestAccessToken(Office.context.mailbox.item.sender.emailAddress);
            }

        });

    };

    function getRestAccessToken(EmailAddress){
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
            if (result.status === "succeeded") {
                var accessToken = result.value;
                getChatMessages(accessToken,EmailAddress);                
            } else {
                // Handle the error
            }
        });
    }
    function getChatMessages(accessToken,emailAddress) {
        var filterString = "SingleValueExtendedProperties/Any(ep: ep/PropertyId eq 'String 0x001a' and ep/Value eq 'IPM.SkypeTeams.Message') and SingleValueExtendedProperties/Any(ep: ep/PropertyId eq 'String 0x5d01' and ep/Value eq '" + emailAddress  + "')";
        var GetURL = "https://outlook.office.com/api/v2.0/me/MailFolders/AllItems/messages?$OrderyBy=ReceivedDateTime desc&$Top=30&$Select=ReceivedDateTime,bodyPreview,webLink&$filter=" + filterString;
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

    function resolveName(NameToLookup){
        var request = GetResolveNameRequest(NameToLookup);
        var EmailAddress = "";        
        Office.context.mailbox.makeEwsRequestAsync(request, function (asyncResult) {
            var parser = new DOMParser();
            var doc = parser.parseFromString(asyncResult.value, "text/xml");
            var values = doc.getElementsByTagName("t:EmailAddress");
            if(values.length != 0){
                EmailAddress = values[0].textContent;
                getRestAccessToken(EmailAddress);
            } 
            

        });

    }
   
    function GetResolveNameRequest(NameToLookup) {
        var results =    

        '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '  <soap:Header>' +
        '    <t:RequestServerVersion Version="Exchange2013" />' +
        '  </soap:Header>' +
        '  <soap:Body>' +
        '    <m:ResolveNames ReturnFullContactData="true" SearchScope="ActiveDirectory">' +
        '      <m:UnresolvedEntry>' + NameToLookup + '</m:UnresolvedEntry>' +
        '    </m:ResolveNames>' +
        '  </soap:Body>' +
        '</soap:Envelope>'
         return results;
    }
    
    function DisplayMessages(Messages) {
        try {
            var html = "<div class=\"ms-Table-row\">";
            html = html + "<span class=\"ms-Table-cell\" >ReceivedDateTime</span>";
            html = html + "<span class=\"ms-Table-cell\">BodyPreview</span>";
            html = html + "</div>";
            Messages.forEach(function (Message) {
                var rcvDate = Date.parse(Message.ReceivedDateTime);
                html = html + "<div class=\"ms-Table-row\">";
                html = html +"<span class=\"ms-Table-cell\">" + rcvDate.toString('dd-MMM-yy HH:mm') + "</span>";
                html = html +"<span id=\"Subject\" class=\"ms-Table-cell\">";
                html = html + Message.BodyPreview + " <a target='_blank' href='" + Message.WebLink + "'> Link</a></span ></div >";
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