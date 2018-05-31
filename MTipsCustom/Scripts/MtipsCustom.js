// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before updating the span elements with values from the current message
Office.initialize = function () {
    $(document).ready(function () {
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

function getMailtips(accessToken) {
     var PostURL = "https://outlook.office.com/api/beta/me/GetMailTips";
     var mtipRequest = "{ \"EmailAddresses\": [ \"gscales@datarumble.com\" ],\"MailTipsOptions\": \"mailboxFullStatus\"}";
    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: PostURL,
        data: mtipRequest,
        dataType: 'json',
        headers: { 'Authorization': 'Bearer ' + accessToken }
    }).done(function (item) {

        var uservalue = item.value[0].EmailAddress.Address;
        $('#UserList').text(uservalue);
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


