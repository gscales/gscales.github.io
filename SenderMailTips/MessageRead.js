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
              getMailTips(accessToken);
          } else {
              // Handle the error
          }
      });

      });

  };

  function getMailTips(accessToken) {
      var PostURL = "https://outlook.office.com/api/beta/me/GetMailTips";
      var mtipRequest = "{ \"EmailAddresses\": [ \"" + Office.context.mailbox.item.sender.emailAddress + "\" ],\"MailTipsOptions\": \"automaticReplies,customMailTip,maxMessageSize,moderationStatus,recipientScope, mailboxFullStatus\"}";
      $.ajax({
          type: "POST",
          contentType: "application/json; charset=utf-8",
          url: PostURL,
          data: mtipRequest,
          dataType: 'json',
          headers: { 'Authorization': 'Bearer ' + accessToken }
      }).done(function (item) {         
          AddMailTipEntry(item.value[0]);         
      }).fail(function (error) {
          // Handle error
      });
  }

  function AddMailTipEntry(entry) {
      var html = `
              <div class="ms-Table-row">
              <span class="ms-Table-cell">Property</span>
              <span class="ms-Table-cell">Value</span>
              </div>
`;
      if (entry.hasOwnProperty("EmailAddress")) {
          var EmailAddress = entry.EmailAddress.Address;
          var html = html + `
                  <div class="ms-Table-row">
                  <span class="ms-Table-cell ms-fontWeight-semibold">EmailAddress</span>
                  <span id="EmailAddress" class="ms-Table-cell">`;
          html = html + EmailAddress + "</span ></div >";
      }
      if (entry.hasOwnProperty("MailboxFull")) {
          var MailboxFullValue = entry.MailboxFull;
          var html = html + `
                  <div class="ms-Table-row">
                  <span class="ms-Table-cell ms-fontWeight-semibold">MailboxFull</span>
                  <span id="MailboxFull" class="ms-Table-cell">`;
          html = html + MailboxFullValue + "</span ></div >";
      }
      if (entry.hasOwnProperty("RecipientScope")) {
          var recipientScope = entry.RecipientScope;
          var html = html + `
                  <div class="ms-Table-row">
                  <span class="ms-Table-cell ms-fontWeight-semibold">RecipientScope</span>
                  <span id="recipientScope" class="ms-Table-cell">`;
          html = html + recipientScope + "</span ></div >";
      }
      if (entry.hasOwnProperty("MaxMessageSize")) {
          var MaxMessageSize = entry.MaxMessageSize;
          var html = html + `
                  <div class="ms-Table-row">
                  <span class="ms-Table-cell ms-fontWeight-semibold">MaxMessageSize</span>
                  <span id="MaxMessageSize" class="ms-Table-cell">`;
          html = html + MaxMessageSize + "</span ></div >";
      }
      if (entry.hasOwnProperty("IsModerated")) {
          var IsModerated = entry.IsModerated;
          var html = html + `
                  <div class="ms-Table-row">
                  <span class="ms-Table-cell ms-fontWeight-semibold">IsModerated</span>
                  <span id="IsModerated" class="ms-Table-cell">`;
          html = html + IsModerated + "</span ></div >";
      }
      if (entry.hasOwnProperty("AutomaticReplies")) {
          if (entry.AutomaticReplies.hasOwnProperty("Message")) {
              var arMessage = entry.AutomaticReplies.Message;
              var html = html + `
                  <div class="ms-Table-row">
                  <span class="ms-Table-cell ms-fontWeight-semibold">AutomaticReply</span>
                  <span id="IsModerated" class="ms-Table-cell">`;
              html = html + arMessage + "</span ></div >";
          }
      }
      if (entry.hasOwnProperty("CustomMailTip")) {
          var CustomMailTip = entry.CustomMailTip;
          var html = html + `
                  <div class="ms-Table-row">
                  <span class="ms-Table-cell ms-fontWeight-semibold">CustomMailTip</span>
                  <span id="IsModerated" class="ms-Table-cell">`;
          html = html + CustomMailTip + "</span ></div >";

      } 
      $('#mTipTable').append(html);
    }

 
  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();