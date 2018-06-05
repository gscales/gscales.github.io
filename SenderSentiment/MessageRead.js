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
                    getSentiment(accessToken);
                } else {
                    // Handle the error
                }
            });

        });

    };

    function getSentiment(accessToken) {
        var rightNow = new Date();
        rightNow.setDate(rightNow.getDate() - 30);
        var GETURL = "https://outlook.office.com/api/v2.0/me/MailFolders/Inbox/messages/?$select=ReceivedDateTime,Sender&$Top=1000&$filter=from/emailAddress/address eq '" + Office.context.mailbox.item.sender.emailAddress + "' And receivedDateTime ge " + rightNow.toISOString() + "&$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String {00062008-0000-0000-C000-000000000046} Name EntityExtraction/Sentiment1.0')";
        GETURL = encodeURI(GETURL);
        $.ajax({
            type: "GET",
            contentType: "application/json; charset=utf-8",
            url: GETURL,
            dataType: 'json',
            headers: { 'Authorization': 'Bearer ' + accessToken }
        }).done(function (item) {
            DisplayResults(item, accessToken);
        }).fail(function (error) {
            // Handle error
        });
    }

    function uuidv4() {
        var uniqueId = Math.random().toString(36).substring(2) + (new Date()).getTime().toString(36);
        return uniqueId;
    }

    function DisplayResults(entry, accessToken) {
        var Negative = 0;
        var Positive = 0;
        var Last7Negative = 0;
        var Last7Positive = 0;
        var Last7Count = 0;
        var last7Date = new Date();
        var BatchId = uuidv4();
        var BatchPost = "";
        BatchPost = BatchPost + "--batch_" + BatchId + "\n";
        var BatchCount = 0;
        last7Date.setDate(last7Date.getDate() - 7);
        var index = 0;
        for (index = 0; index < entry.value.length; ++index) {
            var ParsedDate = new Date(entry.value[index].ReceivedDateTime);
            if (ParsedDate > last7Date) {
                Last7Count++;
            }
            if (entry.value[index].hasOwnProperty("SingleValueExtendedProperties")) {
                BatchPost = BatchPost + "Content-Type:application/http\n";
                BatchPost = BatchPost + "Content-Transfer-Encoding:binary\n\n";
                var GETURL = "GET https://outlook.office.com/api/v2.0/me/MailFolders/Inbox/messages('" + entry.value[index].Id + "')/?$select=ReceivedDateTime,Sender&$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String {00062008-0000-0000-C000-000000000046} Name EntityExtraction/Sentiment1.0')";
                // GETURL = encodeURI(GETURL);
                BatchPost = BatchPost + GETURL + " HTTP/1.1\n\n";
                BatchPost = BatchPost + "--batch_" + BatchId + "\n";
                BatchCount++;
                //To do clean this up to reduce the duplication current POC
                if (BatchCount >= 20) {
                    var headerc = "multipart/mixed; charset=utf-8; boundary=batch_" + BatchId;
                    var PostURL = "https://outlook.office.com/api/v2.0/me/$batch";
                    $.ajax({
                        type: "POST",
                        contentType: headerc,
                        url: PostURL,
                        async: false,
                        data: BatchPost,
                        dataType: 'json',
                        headers: { 'Authorization': 'Bearer ' + accessToken }
                    }).done(function (item) {
                        var resIndex = 0;
                        for (resIndex = 0; resIndex < item.responses.length; ++resIndex) {
                            var ParsedDate = new Date(item.responses[resIndex].body.ReceivedDateTime);
                            var EntityValues = JSON.parse(item.responses[resIndex].body.SingleValueExtendedProperties[0].Value);
                            if (EntityValues[0].hasOwnProperty("sentiment")) {
                                if (EntityValues[0].sentiment.polarity == "positive") {
                                    Positive++;
                                    if (ParsedDate > last7Date) {
                                        Last7Positive++;
                                    }
                                }
                                if (EntityValues[0].sentiment.polarity == "negative") {
                                    Negative++;
                                    if (ParsedDate > last7Date) {
                                        Last7Negative++;
                                    }

                                }
                            }
                        }
                    }).fail(function (error) {
                        //todo handle error
                    });

                    BatchCount = 0
                    BatchId = uuidv4();
                    BatchPost = "";
                    BatchPost = BatchPost + "--batch_" + BatchId + "\n";
                }
            }

        }
        if (BatchCount > 0) {

            var headerc = "multipart/mixed; charset=utf-8; boundary=batch_" + BatchId;
            var PostURL = "https://outlook.office.com/api/v2.0/me/$batch";
            $.ajax({
                type: "POST",
                contentType: headerc,
                url: PostURL,
                async: false,
                data: BatchPost,
                dataType: 'json',
                headers: { 'Authorization': 'Bearer ' + accessToken }
            }).done(function (item) {
                var resIndex = 0;
                for (resIndex = 0; resIndex < item.responses.length; ++resIndex) {
                    var ParsedDate = new Date(item.responses[resIndex].body.ReceivedDateTime);
                    var EntityValues = JSON.parse(item.responses[resIndex].body.SingleValueExtendedProperties[0].Value);
                    if (EntityValues[0].hasOwnProperty("sentiment")) {
                        if (EntityValues[0].sentiment.polarity == "positive") {
                            Positive++;
                            if (ParsedDate > last7Date) {
                                Last7Positive++;
                            }
                        }
                        if (EntityValues[0].sentiment.polarity == "negative") {
                            Negative++;
                            if (ParsedDate > last7Date) {
                                Last7Negative++;
                            }

                        }
                    }
                }
            }).fail(function (error) {
                //todo handle error
            });
        }

        var html = "<div class=\"ms-Table-row\">";
        html = html + "<span class=\"ms-Table-cell\"></span>";
        html = html + "<span class=\"ms-Table-cell\"></span>";
        html = html + "</div>";

        html = html + "<div class=\"ms-Table-row\">";
        html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\">TotalCount Last 30 Days</span>";
        html = html + "<span id=\"last30\" class=\"ms-Table-cell\">";
        html = html + entry.value.length + "</span ></div >";
        html = html + "<div class=\"ms-Table-row\">";
        html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\">TotalCount Last 7 Days</span>";
        html = html + "<span id=\"last7\" class=\"ms-Table-cell\">";
        html = html + Last7Count + "</span ></div >";
        $('#sentTable').append(html);
        var html = "";
        html = html + "<div class=\"ms-Table-row\">";
        html = html + "<span class=\"ms-Table-cell\"></span>";
        html = html + "<span class=\"ms-Table-cell\"  style=\"font-size: 14px;\">Positive &#x1f603;</span>";
        html = html + "<span class=\"ms-Table-cell\"  style=\"font-size: 14px;\">Negative &#x1f641;</span>";
        html = html + "</div>";

        html = html + "<div class=\"ms-Table-row\">";
        html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\">Last 7 Days</span>";
        html = html + "<span id=\"pCount7\" class=\"ms-Table-cell\" style=\"text-align:center;\">";
        html = html + Last7Positive + "</span >";
        html = html + "<span id=\"nCount7\" class=\"ms-Table-cell\" style=\"text-align:center;\">" + Last7Negative + "</span ></div >";
        html = html + "<div class=\"ms-Table-row\">";
        html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\">Last 30 Days</span>";
        html = html + "<span id=\"pCount\" class=\"ms-Table-cell\" style=\"text-align:center;\">";
        html = html + Positive + "</span >";
        html = html + "<span id=\"nCount\" class=\"ms-Table-cell\" style=\"text-align:center;\">" + Negative + "</span ></div >";
        $('#sentimentTable').append(html);
    }
    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();