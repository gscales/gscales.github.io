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
                getFolderIdFromProperty();
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

    function getFolderId(accessToken){
        var GetURL = "https://outlook.office.com/api/v2.0/me/MailFolders/Inbox/" + filterString;
        $.ajax({
            type: "Get",
            contentType: "application/json; charset=utf-8",
            url: GetURL,
            dataType: 'json',
            headers: { 'Authorization': 'Bearer ' + accessToken }
        }).done(function (item) {
            console.log(item.value);
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

    function getFolderIdFromProperty(){
        var request = GetChatMessagesFolderIdRequest();
  
        Office.context.mailbox.makeEwsRequestAsync(request, function (asyncResult) {
            var parser = new DOMParser();
            var doc = parser.parseFromString(asyncResult.value, "text/xml");
            var exProp = doc.getElementsByTagName("t:ExtendedProperty");
            if(exProp.length != 0){
                ConvertEWSId(base64ToHex(exProp[0].textContent));
            }        

        });

    }
    function ConvertEWSId(IdToConvert){
        var request = ConvertIdRequest(IdToConvert);
     
        Office.context.mailbox.makeEwsRequestAsync(request, function (asyncResult) {
            var parser = new DOMParser();
            var doc = parser.parseFromString(asyncResult.value, "text/xml");
            var aId = doc.getElementsByTagName("m:AlternateId");
            if(aId.length != 0){
               console.log(aId);
            }        

        });

    }

    function base64ToHex(str) {
        const raw = atob(str);
        let result = '';
        for (let i = 0; i < raw.length; i++) {
          const hex = raw.charCodeAt(i).toString(16);
          result += (hex.length === 2 ? hex : '0' + hex);
        }
        return result.toUpperCase();
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

    function GetChatMessagesFolderIdRequest(){
        var RequestString =    

        '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '  <soap:Header>' +
        '    <t:RequestServerVersion Version="Exchange2016" />' +
        '  </soap:Header>' +
        '  <soap:Body>' +
        '<m:GetFolder>' +
        '<m:FolderShape>' +
        '  <t:BaseShape>AllProperties</t:BaseShape>' +
        '  <t:AdditionalProperties>' +
        '    <t:ExtendedFieldURI PropertySetId="e49d64da-9f3b-41ac-9684-c6e01f30cdfa" PropertyName="TeamsMessagesDataFolderEntryId" PropertyType="Binary" />' +
        '  </t:AdditionalProperties>' +
        '</m:FolderShape>' +
        '<m:FolderIds>' +
        '   <t:DistinguishedFolderId Id="inbox">' +
        '  </t:DistinguishedFolderId>' +
        '</m:FolderIds>' +
        '</m:GetFolder>' +
        '  </soap:Body>' +
        '</soap:Envelope>'
         return RequestString;
    }
    

    function ConvertIdRequest(IdToConvert){
        var RequestString =  

        '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '  <soap:Header>' +
        '    <t:RequestServerVersion Version="Exchange2016" />' +
        '  </soap:Header>' +
        '  <soap:Body>' +
        '<m:ConvertId DestinationFormat="EwsId">' +
        '<m:SourceIds>' +
        '  <t:AlternateId Format="HexEntryId" Id="' + IdToConvert + '" Mailbox="blah@blah.com" />'+
        ' </m:SourceIds>' +
        '</m:ConvertId>' +
        '  </soap:Body>' +
        '</soap:Envelope>'
        return RequestString;
  
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