
(function () {

    const GetTokenFromSkype = (TeamsClient) => {
        return new Promise(
            (resolve, reject) => {

            }
        );
    }


    async function authtest() {
        try {
            return "test12345";
            
        }
        catch (error) {
            //handle Error **To Do **
        }
    }

    if (typeof module !== 'undefined' && typeof module.exports !== 'undefined')
        module.exports.authtest = function () {
            return authtest();
        }
    else {
        window.authtest = authtest;
    }
       



}());

