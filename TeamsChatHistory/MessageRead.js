    var messageBanner;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            ExecuteSearch();

        });

    };

    function ExecuteSearch(){
        if(Office.context.mailbox.item.sender.emailAddress == "noreply@email.teams.microsoft.com"){
            resolveName(Office.context.mailbox.item.sender.displayName.replace(" in Teams",""));
        }else{
            getTeamsMessagesFolder(Office.context.mailbox.item.sender.emailAddress);
        }
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
                getTeamsMessagesFolder(EmailAddress);
            }        

        });

    }

    function getTeamsMessagesFolder(EmailAddress){
        var request = FindFolderRequest();
  
        Office.context.mailbox.makeEwsRequestAsync(request, function (asyncResult) {
            var parser = new DOMParser();
            var doc = parser.parseFromString(asyncResult.value, "text/xml");
            var folderid = doc.getElementsByTagName("t:FolderId");
            if(folderid.length != 0){
                FindItems(folderid[0].getAttribute('Id'),EmailAddress);
            }        

        });

    }

    function FindItems(FolderId,EmailAddress){
        var request = FindItemsRequest(FolderId,EmailAddress);  
        Office.context.mailbox.makeEwsRequestAsync(request, function (asyncResult) {
            var parser = new DOMParser();
            var doc = parser.parseFromString(asyncResult.value, "text/xml");
            var Items = doc.getElementsByTagName("t:Message");
            DisplayMessages(Items);

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

    function FindItemsRequest(FolderId,EmailAddress) {
        var DaysToSubtrace = $( "#lbts" ).val();
        var StartDate = new Date();
        StartDate.setDate(StartDate.getDate()-DaysToSubtrace);
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
          '   <t:IsEqualTo>' +
          '   <t:ExtendedFieldURI PropertyTag="23809" PropertyType="String" />' +
          '    <t:FieldURIOrConstant>' +
          '     <t:Constant Value="' + EmailAddress + '" />' +
          '   </t:FieldURIOrConstant>' +
          '   </t:IsEqualTo>' +
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
            $('#mTchatTable').empty().append(html);
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
