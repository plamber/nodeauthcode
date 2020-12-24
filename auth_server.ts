import { AuthenticationContext, ErrorResponse, TokenResponse } from 'adal-node';
const http = require('http');
const url = require('url');

// create an azure ad client with redirect URL http://localhost. Assign graph delegated permissions and create a secret. Put the values below.
const clientId = "put your client id here"
const tenantID = "put your tenant id here"
// use null and you get an exception
const secret = "put your secret here";

const resource = "https://graph.microsoft.com"
const authorityUrl = `https://login.microsoftonline.com/${tenantID}`
const open = require('open');
let redirectUri = "";

const authCtx: AuthenticationContext = new AuthenticationContext(authorityUrl);

// Create http server.
var httpServer = http.createServer(function (request : any, response : any) {
    const outputError = (response, intro, message): void => {
        response.writeHead(200, {'Access-Control-Allow-Origin':'*','Content-Type': 'text/plain'});
        response.write("The service replied with following error message.");
        if (intro !== undefined) {
            response.write("\n")
            response.write(intro);
        }
        if (message !== undefined) {
            response.write("\n")
            response.write(message);
        }      
        response.end();
    }

    const outputSuccess = (response, token: TokenResponse): void => {
        response.writeHead(200, {'Access-Control-Allow-Origin':'*','Content-Type': 'text/plain'});
        response.write("The service returns following information");
        response.write(`\n\nAccess token: ${token.accessToken}`);
        response.write(`\n\nRefresh token: ${token.refreshToken}`);
        response.write(`\n\nUser ID: ${token.userId}`);
        response.end();
    }
    
    const queryString = url.parse(request.url, true).query;
    if (queryString.error !== undefined) {
        outputError(response, queryString.error, queryString.error_description);
    }
    if (queryString.code !== undefined) {
        const authorizationCode = queryString.code;
        authCtx.acquireTokenWithAuthorizationCode(authorizationCode, redirectUri, resource, clientId, secret,
        (error: Error, rsp: TokenResponse | ErrorResponse): void => {
            if (error) {
              outputError(response, rsp.error, (rsp as any).error_description);
              return;
            }
            outputSuccess(response, rsp as TokenResponse)
          });
    }
});

// pick a random port
httpServer.listen(0, () => {
    redirectUri = `http://localhost:${httpServer.address().port}`
    const requestState = Math.random().toString(16).substr(2, 20);
    const urlToOpen = `${authorityUrl}/oauth2/authorize?response_type=code&client_id=${clientId}&redirect_uri=${redirectUri}&state=${requestState}&resource=${resource}&prompt=select_account`
    open(urlToOpen);
});

