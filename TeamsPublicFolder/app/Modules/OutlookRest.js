const GetEmailAddress = (idToken) => {
    return new Promise(
        (resolve, reject) => {            
            $.ajax({
                type: "GET",
                contentType: "application/json; charset=utf-8",
                url: ("https://outlook.office.com/api/v2.0/me"),
                dataType: 'json',
                headers: { 'Authorization': 'Bearer ' + idToken }
            }).done(function (item) {
                resolve(item.EmailAddress);
            }).fail(function (error) {
                reject(error);
            });
        }
    );
}