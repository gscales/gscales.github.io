const GetGroupMembers = (idToken, teamscontext) => {
    return new Promise(
        (resolve, reject) => {
            GroupId = teamscontext.groupId;
            $.ajax({
                type: "GET",
                contentType: "application/json; charset=utf-8",
                url: ("https://graph.microsoft.com/v1.0/groups/" + GroupId + "/members"),
                dataType: 'json',
                headers: { 'Authorization': 'Bearer ' + idToken }
            }).done(function (item) {
                resolve(item);
            }).fail(function (error) {
                reject(error);
            });
        }
    );
}

const DownloadOneDriveFile = (Token,URL) => {
    return new Promise(
        (resolve, reject) => {

            fetch(URL, {
                credentials: 'Bearer'
              });
        }
    );
}
const GenericGraphGet = (Token,URL) => {
    return new Promise(
        (resolve, reject) => {
            $.ajax({
                type: "GET",
                contentType: "application/json; charset=utf-8",
                url: URL,
                dataType: 'json',
                headers: { 'Authorization': 'Bearer ' + Token }
            }).done(function (item) {
                resolve(item);
            }).fail(function (error) {
                var location = error.getResponseHeader("Location");
                if (location !== null) {
                    resolve(location);
                }
                reject(error);
            });
        }
    );
}
const GenericGraphPOST = (Token, URL, POSTData) => {
    return new Promise(
        (resolve, reject) => {
            $.ajax({
                type: "POST",
                contentType: "application/json; charset=utf-8",
                url: URL,
                dataType: 'json',
                data: POSTData,
                headers: { 'Authorization': 'Bearer ' + Token }
            }).done(function (item) {
                resolve(item);
            }).fail(function (error) {
                reject(error);
            });
        }
    );
};

const WorkBookPOST = (Token, URL, SessionId ,POSTData) => {
    return new Promise(
        (resolve, reject) => {
            var PostData = {};
            $.ajax({
                type: "POST",
                contentType: "application/json; charset=utf-8",
                url: URL,
                dataType: 'json',
                data: POSTData,
                headers: {
                    'Authorization': 'Bearer ' + Token,
                    'workbook-session-id': SessionId
                }
            }).done(function (item) {
                resolve(item);
            }).fail(function (error) {
                reject(error);
            });
        }
    );
};

const WorkBookPATCH = (Token, URL, SessionId ,POSTData) => {
    return new Promise(
        (resolve, reject) => {
            var PostData = {};
            $.ajax({
                type: "PATCH",
                contentType: "application/json; charset=utf-8",
                url: URL,
                dataType: 'json',
                data: POSTData,
                headers: {
                    'Authorization': 'Bearer ' + Token,
                    'workbook-session-id': SessionId
                }
            }).done(function (item) {
                resolve(item);
            }).fail(function (error) {
                reject(error);
            });
        }
    );
};
const CreateOneDriveFile = (Token, URL, fileData) => {
    return new Promise(
        (resolve, reject) => { 
            $.ajax({
                type: "PUT",
                contentType: "application/octet-stream",
                url: URL,
                data: fileData,
                processData: false,
                headers: { 'Authorization': 'Bearer ' + Token }
            }).done(function (item) {
                resolve(item);
            }).fail(function (error) {
                reject(error);
            });
        }
    );
};


function formatDate(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;

    return [year, month, day].join('-');
}

function addDays(date, days) {
    var result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
}







