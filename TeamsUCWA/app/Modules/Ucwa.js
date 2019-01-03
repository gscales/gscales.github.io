const DiscoverUCWAEndpointStage1 = (DomainName) => {
    return new Promise(
        (resolve, reject) => {
            var DiscoverURL = 'https://webdir.online.lync.com/autodiscover/autodiscoverservice.svc/root' + "?originalDomain=" + DomainName;              
            $.getJSON(DiscoverURL, function(data, status){
                resolve(data);
            }).error(function() { reject("error"); });
        }
    );
}
const DiscoverUCWAEndpointStage1Redirect = (RedirectURL) => {
    return new Promise(
        (resolve, reject) => {          
            $.getJSON(RedirectURL, function(data, status){
                resolve(data);
            }).error(function() { reject("error"); });
        }
    );
}

const DiscoverUCWAEndpointStage2 = (URL,Token) => {
    return new Promise(
        (resolve, reject) => {     
            $.ajax({
                type: "GET",
                url: URL,
                contentType: JSON,           
                beforeSend: function (xhr) {   //Set token here
                    xhr.setRequestHeader("Authorization", 'Bearer '+ Token);
                }
            }).done(function (response) {
                resolve(response);
            }).fail(function (err)  {
                reject(err);
            });

        }
    );
}

const ConnectUCWA = (URL,Token) => {
    return new Promise(
        (resolve, reject) => {  
            $UCWAEndpoint = {};
            $UCWAEndpoint.UserAgent = "TeamsUCWA";
            $UCWAEndpoint.Culture = "en-US";
            $UCWAEndpoint.EndpointId = uuidv4();
            $.ajax({
                type: "POST",
                url: URL,
                contentType: 'application/json; charset=utf-8',  
                data: JSON.stringify($UCWAEndpoint),         
                beforeSend: function (xhr) {   //Set token here
                    xhr.setRequestHeader('X-MS-RequiresMinResourceVersion', '2');
                    xhr.setRequestHeader('X-Ms-Namespace', 'internal');
                    xhr.setRequestHeader("Authorization", 'Bearer '+ Token);
                }
            }).done(function (response) {
                resolve(response);
            }).fail(function (err)  {
                reject(err);
            });

        }
    );
}


const GetCommunicationLinks = (URL,Token) => {
    return new Promise(
        (resolve, reject) => {  
            $.ajax({
                type: "GET",
                url: URL,
                contentType: 'application/json; charset=utf-8',    
                beforeSend: function (xhr) {   //Set token here
                    xhr.setRequestHeader("Authorization", 'Bearer '+ Token);
                },
                success : function(data, status,jqXHR ) { 
                    resolve(data);
                },
            }).fail(function (err)  {
                reject(err);
            });

        }
    );
}

const EnableConversationHistory = (URL,Token,Data) => {
    return new Promise(
        (resolve, reject) => {  
            var matchheader = "\"" +  Data.etag + "\"";
            $.ajax({
                type: "PUT",
                url: URL,
                contentType: 'application/json; charset=utf-8',  
                data: JSON.stringify(Data),         
                beforeSend: function (xhr) {   //Set token here
                    xhr.setRequestHeader("Authorization", 'Bearer '+ Token);
                    xhr.setRequestHeader('If-Match', matchheader)
                }
            }).done(function (response) {
                resolve(response);
            }).fail(function (err)  {
                reject(err);
            });

        }
    );
}

const GetEvents = (URL,Token) => {
    return new Promise(
        (resolve, reject) => {  
            $.ajax({
                type: "GET",
                url: (URL + "&timeout=60"),
                contentType: 'application/json; charset=utf-8',    
                beforeSend: function (xhr) {   //Set token here
                    xhr.setRequestHeader("Authorization", 'Bearer '+ Token);
                },
                success : function(data, status,jqXHR ) { 
                    resolve(data);
                },
            }).fail(function (err)  {
                reject(err);
            });

        }
    );
}

const GetConversationLogs = (URL,Token) => {
    return new Promise(
        (resolve, reject) => {  
            $.ajax({
                type: "GET",
                url: URL,
                contentType: 'application/json; charset=utf-8',    
                beforeSend: function (xhr) {   //Set token here
                    xhr.setRequestHeader("Authorization", 'Bearer '+ Token);
                },
                success : function(data, status,jqXHR ) { 
                    resolve(data);
                },
            }).fail(function (err)  {
                reject(err);
            });

        }
    );
}

const BatchConversationPost = (ServerHostName, URL,postdata,BatchId,Token) => {
    return new Promise(
        (resolve, reject) => {  
            
            $.ajax({
                type: "POST",
                url: URL,
                contentType: ('multipart/batching;boundary=' + BatchId), 
                data: postdata,   
                beforeSend: function (xhr) {   //Set token here
                    xhr.setRequestHeader("Authorization", 'Bearer '+ Token);
                    xhr.setRequestHeader("Accept",'multipart/batching');
                },
                success : function(data, status,jqXHR ) {  
                    let contentTypeHeader = jqXHR.getResponseHeader("Content-Type");
                    let RBatchId = contentTypeHeader.split(";")[1].split("=")[1]
                    RBatchId = RBatchId.substr(1,RBatchId.length-2);
                    //console.log(RBatchId);
                    let responseLines = data.split('--' + RBatchId);
                    var JsonIndex = [];
                    var vnd = 0;
                    responseLines.forEach(function (response) {                        
                        var startJson = response.indexOf('{');
                        var endJson = response.lastIndexOf('}');
                        if (startJson < 0 || endJson < 0) {
                            return;
                        }
                        try{
                            var responseJson = JSON.parse(response.substr(startJson, (endJson - startJson) + 1));
                            JsonIndex.push(responseJson);
                        }catch(error){
                            console.log(error);
                        }
                        
                        
                    });                   
                    resolve(JsonIndex);
                },
            }).fail(function (err)  {
                reject(err);
            });

        }
    );
}

