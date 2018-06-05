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
      var html = "<div class=\"ms-Table-row\">";
      html = html + "<span class=\"ms-Table-cell\">Property</span>";
      html = html + "<span class=\"ms-Table-cell\">Value</span>";
      html = html + "</div>";
      if (entry.hasOwnProperty("EmailAddress")) {
          var EmailAddress = entry.EmailAddress.Address;
          html = html + "<div class=\"ms-Table-row\">";
          html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\">EmailAddress</span>";
          html = html + "<span id=\"EmailAddress\" class=\"ms-Table-cell\">";
          html = html + EmailAddress + "</span ></div >";
      }
      if (entry.hasOwnProperty("MailboxFull")) {
          var MailboxFullValue = entry.MailboxFull;
          html = html + "<div class=\"ms-Table-row\">";
          html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\">MailboxFull</span>";
          html = html + "<span id=\"MailboxFull\" class=\"ms-Table-cell\">";
          html = html + MailboxFullValue + "</span ></div >";
      }
      if (entry.hasOwnProperty("RecipientScope")) {
          var recipientScope = entry.RecipientScope;
          html = html + "<div class=\"ms-Table-row\">";
          html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\">RecipientScope</span>";
          html = html + "<span id=\"RecipientScope\" class=\"ms-Table-cell\">";
          html = html + recipientScope + "</span ></div >";
      }
      if (entry.hasOwnProperty("MaxMessageSize")) {
          var MaxMessageSize = entry.MaxMessageSize;
          html = html + "<div class=\"ms-Table-row\">";
          html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\">MaxMessageSize</span>";
          html = html + "<span id=\"MaxMessageSize\" class=\"ms-Table-cell\">";
          html = html + MaxMessageSize + "</span ></div >";
      }
      if (entry.hasOwnProperty("IsModerated")) {
          var IsModerated = entry.IsModerated;
          html = html + "<div class=\"ms-Table-row\">";
          html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\">IsModerated</span>";
          html = html + "<span id=\"IsModerated\" class=\"ms-Table-cell\">";
          html = html + IsModerated + "</span ></div >";
      }
      if (entry.hasOwnProperty("AutomaticReplies")) {
          if (entry.AutomaticReplies.hasOwnProperty("Message")) {
              var arMessage = entry.AutomaticReplies.Message;
              html = html + "<div class=\"ms-Table-row\">";
              html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\">AutomaticReply</span>";
              html = html + "<span id=\"AutomaticReply\" class=\"ms-Table-cell\">";
              html = html + arMessage + "</span ></div >";
          }
      }
      if (entry.hasOwnProperty("CustomMailTip")) {
          var CustomMailTip = entry.CustomMailTip;
          html = html + "<div class=\"ms-Table-row\">";
          html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\">CustomMailTip</span>";
          html = html + "<span id=\"CustomMailTip\" class=\"ms-Table-cell\">";
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