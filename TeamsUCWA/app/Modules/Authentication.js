const Authuser = (upn,appconfig,url,hostname) => {
    return new Promise(
        (resolve, reject) => {
            let resourceURL = "https://" + hostname;
            let config = {
                clientId: appConfig.clientId,
                redirectUri: window.location.origin + appConfig.redirectUri,       // This should be in the list of redirect uris for the AAD app
                cacheLocation: "localStorage",
                navigateToLoginRequestUrl: false,              
                endpoints: {
                    resourceURL: resourceURL
                }
            };
            if (upn) {
                config.extraQueryParameters = "scope=openid+profile&login_hint=" + encodeURIComponent(upn) + "";
            } else {
                config.extraQueryParameters = "scope=openid+profile";
            }           
            let authContext = new AuthenticationContext(config);
            microsoftTeams.authentication.authenticate({
                url: window.location.origin + appConfig.authwindow, 
                width: 400,
                height: 400,
                successCallback: function (t) {
                    // Note: token is only good for one hour
                    token = t;
                    resolve(token);
                },
                failureCallback: function (err) {
                      reject(err);
                }
            });
        }
        );
}


