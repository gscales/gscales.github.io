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
        var request = FindFolderRequest();
  
        Office.context.mailbox.makeEwsRequestAsync(request, function (asyncResult) {
            var parser = new DOMParser();
            var doc = parser.parseFromString(asyncResult.value, "text/xml");
            var folderid = doc.getElementsByTagName("t:FolderId");
            if(folderid.length != 0){
                FindItems(folderid[0].getAttribute('Id'));
            }        

        });

    }

    function FindItems(FolderId){
        var request = FindItemsRequest(FolderId);  
        Office.context.mailbox.makeEwsRequestAsync(request, function (asyncResult) {
            var parser = new DOMParser();
            var doc = parser.parseFromString(asyncResult.value, "text/xml");
            var Items = doc.getElementsByTagName("t:Message");
            DisplayMessages(Items);

        });
    }
    function ConvertEWSId(IdToConvert){
        var ConvertIdRequestString = ConvertIdRequest(IdToConvert);
     
        Office.context.mailbox.makeEwsRequestAsync(ConvertIdRequestString, function (asyncResult) {
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
        '  <t:AlternateId Format="HexEntryId" Id="' + IdToConvert + '" Mailbox="' + Office.context.mailbox.userProfile.emailAddress + '" />'+
        ' </m:SourceIds>' +
        '</m:ConvertId>' +
        '  </soap:Body>' +
        '</soap:Envelope>'
        return RequestString;
  
    }

    function FindFolderRequest(){
        var RequestString =
        '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '  <soap:Header>' +
        '    <t:RequestServerVersion Version="Exchange2016" />' +
        '  </soap:Header>' +
        '  <soap:Body>' +
        '<m:FindFolder Traversal="Shallow">' +
        '<m:FolderShape>' +
        '  <t:BaseShape>AllProperties</t:BaseShape>' +
        '</m:FolderShape>' +
        '<m:IndexedPageFolderView MaxEntriesReturned="1" Offset="0" BasePoint="Beginning" />' +
        '<m:Restriction>' +
        '  <t:IsEqualTo>' +
        '    <t:FieldURI FieldURI="folder:DisplayName" />' +
        '    <t:FieldURIOrConstant>' +
        '      <t:Constant Value="TeamsMessagesData" />' +
        '    </t:FieldURIOrConstant>' +
        '  </t:IsEqualTo>' +
        '</m:Restriction>' +
        '<m:ParentFolderIds>' +
        '  <t:DistinguishedFolderId Id="root" />' +
        '</m:ParentFolderIds>' +
        '</m:FindFolder>' +
        '  </soap:Body>' +
        '</soap:Envelope>'
        return RequestString;
  
    }

    function FindItemsRequest(FolderId) {
        var StartDate = new Date();
        StartDate.setMonth(StartDate.getMonth() - 1);
        var EndDate = new Date();
        var RequestString =
          '<?xml version="1.0" encoding="utf-8"?>' +
          '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
          '  <soap:Header>' +
          '    <t:RequestServerVersion Version="Exchange2016" />' +
          '  </soap:Header>' +
          '  <soap:Body>' +
          '<m:FindItem Traversal="Shallow">' +
          '<m:ItemShape>' +
          '  <t:BaseShape>IdOnly</t:BaseShape>' +
          '  <t:AdditionalProperties>' +  
          '  <t:FieldURI FieldURI="item:Preview" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" />' +
          '  <t:FieldURI FieldURI="item:DateTimeReceived" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" />' +
          '  <t:FieldURI FieldURI="item:WebClientReadFormQueryString" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" />' +          
          '  </t:AdditionalProperties>' +  
          '</m:ItemShape>' +
          '<m:IndexedPageItemView MaxEntriesReturned="1000" Offset="0" BasePoint="Beginning" />' +
          '<m:Restriction>' +
          '  <t:And>' +
          '    <t:IsGreaterThan>' +
          '      <t:FieldURI FieldURI="item:DateTimeReceived" />' +
          '      <t:FieldURIOrConstant>' +
          '        <t:Constant Value="' + StartDate.toISOString() + '" />' +
          '      </t:FieldURIOrConstant>' +
          '    </t:IsGreaterThan>' +
          '    <t:IsLessThan>' +
          '      <t:FieldURI FieldURI="item:DateTimeReceived" />' +
          '      <t:FieldURIOrConstant>' +
          '        <t:Constant Value="' + EndDate.toISOString() + '" />' +
          '      </t:FieldURIOrConstant>' +
          '    </t:IsLessThan>' +
          ' </t:And>' +
          '</m:Restriction>' +
          '<m:ParentFolderIds>' +
          ' <t:FolderId Id="' + FolderId + '" />' +
          '</m:ParentFolderIds>' +
          '</m:FindItem>' +
          '  </soap:Body>' +
          '</soap:Envelope>'
        return RequestString;
      }
    function DisplayMessages(Messages) {
        try {
            var html = "<div class=\"ms-Table-row\">";
            html = html + "<span class=\"ms-Table-cell\">ReceivedDateTime</span>";
            html = html + "<span class=\"ms-Table-cell\">BodyPreview</span>";
            html = html + "</div>";
            for (let Message of Messages) {              
                var rcvDate = new Date(Message.childNodes[1].textContent);
                html = html + "<div class=\"ms-Table-row\">";
                html = html +"<span class=\"ms-Table-cell\">" + rcvDate.toString('dd-MMM-yy HH:mm') + "</span>";
                html = html +"<span id=\"Subject\" class=\"ms-Table-cell\">";
                html = html + Message.childNodes[3].textContent + " <a target='_blank' href='" +  Message.childNodes[2].textContent + "'> Link</a></span ></div >";
            }
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