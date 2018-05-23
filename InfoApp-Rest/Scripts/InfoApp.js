// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before updating the span elements with values from the current message
Office.initialize = function () {
    $(document).ready(function () {
        var item = Office.context.mailbox.userProfile;
        var diag = Office.context.mailbox.diagnostics;
        $('#DisplayName').text("DisplayName : " + item.displayName);
        $('#EmailAddress').text("EmailAddress : " + item.emailAddress);
        $('#TimeZone').text("Timezone : " + item.timeZone);
        $('#displayLanguage').text("Display Language : " + Office.context.displayLanguage);
        $('#HostVersion').text("HostVersion : " + diag.hostVersion);
        $('#UserAgent').text("UserAgent : " + window.navigator.userAgent);
        $('#OWAView').text("OWAView : " + Office.context.mailbox.ewsUrl);
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
            if (result.status === "succeeded") {
                var accessToken = result.value;
                getCurrentItem(accessToken);
            } else {
                // Handle the error
            }
        });
    });
};

function getCurrentItem(accessToken) {
    // Get the item's REST ID
    var itemId = getItemRestId();

    // Construct the REST URL to the current item
    // Details for formatting the URL can be found at 
    // https://msdn.microsoft.com/office/office365/APi/mail-rest-operations#get-a-message-rest
    //var getMessageUrl = Office.context.mailbox.restUrl +
    //    '/v2.0/me/messages/' + itemId;
    var getMessageUrl = "https://graph.microsoft.com/beta/me/messages" + itemId;
    $.ajax({
        url: getMessageUrl,
        dataType: 'json',
        headers: { 'Authorization': 'Bearer ' + accessToken }
    }).done(function (item) {
        // Message is passed in `item`
        var subject = item.Subject;
        $('#ServerName').text("Message Subject : " + subject);
    }).fail(function (error) {
        // Handle error
    });
}

function getItemRestId() {
    if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
        // itemId is already REST-formatted
        return Office.context.mailbox.item.itemId;
    } else {
        // Convert to an item ID for API v2.0
        return Office.context.mailbox.convertToRestId(
            Office.context.mailbox.item.itemId,
            Office.MailboxEnums.RestVersion.v2_0
        );
    }
}


