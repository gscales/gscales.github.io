const GetPublicFolderItems = (idToken, appConfig,Query,mailbox) => {
    return new Promise(
        (resolve, reject) => {
            var GetItemPost = {};
            GetItemPost.hasFolderId = false;
			GetItemPost.Mailbox = mailbox;
            GetItemPost.FolderPath = appConfig.folderpath;
            if(Query != null){
                GetItemPost.Query = Query; 
            }
			GetItemPost.Offset =  0;
			GetItemPost.PageCount = 100;
            var gipRequest = JSON.stringify(GetItemPost);
            $.ajax({
                type: "POST",
                contentType: "application/json; charset=utf-8",
                url: appConfig.ewsproxy,
                dataType: 'json',
                data: gipRequest,
                headers: {
                    'Authorization': 'Bearer ' + idToken,                  

                }
            }).done(function (items) {
                resolve(items);
            }).fail(function (error) {
                reject(error);
            });
        }
    );
}

const RefreshPublicFolderItems = (idToken, CurrentResults,Query,appConfig,offset) => {
    return new Promise(
        (resolve, reject) => {
            var GetItemPost = {};
            GetItemPost.hasFolderId = true;
			GetItemPost.UniqueId = CurrentResults.FolderId.UniqueId;
            GetItemPost.RoutingHeader = CurrentResults.RoutingHeader;
            if(offset){
                GetItemPost.Offset = CurrentResults.nextPageOffset;
            }else{
                GetItemPost.Offset =  0;
            }           
            if(Query != null){
                GetItemPost.Query = Query; 
            }
			GetItemPost.PageCount = 100;
            var gipRequest = JSON.stringify(GetItemPost);
            $.ajax({
                type: "POST",
                contentType: "application/json; charset=utf-8",
                url: appConfig.ewsproxy,
                dataType: 'json',
                data: gipRequest,
                headers: {
                    'Authorization': 'Bearer ' + idToken,                  

                }
            }).done(function (items) {
                resolve(items);
            }).fail(function (error) {
                reject(error);
            });
        }
    );
}











