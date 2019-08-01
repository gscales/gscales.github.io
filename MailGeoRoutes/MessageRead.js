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
              getMailHeaders(accessToken);
          } else {
              // Handle the error
          }
      });

      });

  };

  function getMailHeaders(accessToken) {
      var GetURL = "https://outlook.office.com/api/v2.0/me/messages/" + Office.context.mailbox.item.itemId +
      "?$select=SingleValueExtendedProperties&$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String 0x007D')";   
      $.ajax({
          type: "Get",
          contentType: "application/json; charset=utf-8",
          url: GetURL,
          dataType: 'json',
          headers: { 'Authorization': 'Bearer ' + accessToken }
      }).done(function (item) {         
        ParseMessageHeader(item.SingleValueExtendedProperties[0].Value);       
      }).fail(function (error) {
          // Handle error
      });
  }

function ParseMessageHeader(Header){
  var RegExIP = /\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b/g;
  var result = Header.match(RegExIP);
  var UniqueMatches =  [...new Set(result)];
  var html = "<div class=\"ms-Table-row\">";
  html = html + "<span class=\"ms-Table-cell\">IPAddress</span>";
  html = html + "<span class=\"ms-Table-cell\">Location</span>";
  html = html + "<span class=\"ms-Table-cell\">Organization</span>";
  html = html + "</div>";
  $('#MailRoutesTable').append(html);
  UniqueMatches.forEach(function(Message){    
    LookupIP(Message);
  });
  
  console.log("Done");
}
 
function LookupIP(IPValue){
  var GetURL = "https://api.ip.sb/geoip/" + IPValue;
    $.ajax({
      type: "Get",
      url: GetURL
  }).done(function (item) {     
    
      var ReturnValue = "";
      if(item.hasOwnProperty("country")){
        ReturnValue = item.country;
      }else{
        if(item.hasOwnProperty("timezone")){
          ReturnValue = item.timezone;
        }
      }
      if(item.hasOwnProperty("region")){
        ReturnValue += " " + item.region;
      }
      if(item.hasOwnProperty("city")){
        ReturnValue += " " + item.city;
      }
      var OrganizationVaule = "";
      if(item.hasOwnProperty("organization")){
        OrganizationVaule = item.organization
      }      
      var html = "<div class=\"ms-Table-row\">";
      html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\">" + IPValue + "</span>";
      html = html + "<span id=\"IPAddressVal\" class=\"ms-Table-cell\">";
      html = html + ReturnValue + "</span >";
      html = html + "<span id=\"org\" class=\"ms-Table-cell\">";
      html = html +  OrganizationVaule + "</span ></div >";
      $('#MailRoutesTable').append(html);
  }).fail(function (error) {
      // Handle error
  });
}

  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();


