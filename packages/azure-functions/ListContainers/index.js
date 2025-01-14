const msal = require('@azure/msal-node');
const jwt = require('jsonwebtoken')
const jwksClient = require('jwks-rsa');
require('isomorphic-fetch'); // required for graph library and SharePoint calls
const Graph = require('@microsoft/microsoft-graph-client');
//const { getGraphToken } = require('../utils/auth.js');

// const { getGraphToken } = require('../utils/auth.js');



// const DISCOVERY_KEYS_ENDPOINT = process.env["DISCOVERY_KEYS_ENDPOINT"];
// const config = {
//     auth: {
//         clientId: process.env["APP_CLIENT_ID"],
//         authority: process.env["APP_AUTHORITY"],
//         audience: process.env["APP_AUDIENCE"],
//         clientSecret: process.env["APP_CLIENT_SECRET"]
//     },
//     system: {
//         loggerOptions: {
//             loggerCallback(loglevel, message, containsPii) {
//                 console.log(message);
//             },
//             piiLoggingEnabled: false,
//             logLevel: msal.LogLevel.Verbose,
//         }
//     }
// };
// const cca = new msal.ConfidentialClientApplication(config);

// const isJwtValid = (token) => {
//     if (!token) {
//         return false;
//     }
//     const validationOptions = {
//         algorithms: ['RS256'],
//         audience: config.auth.audience, // v2.0 token
//         issuer: config.auth.issuer // v2.0 token
//         // Also verify JWT has the Container.Manage scope 
//     }
//     jwt.verify(token, getSigningKeys, validationOptions, (err, payload) => {
//         if (err) {
//             console.log(err);
//             return false;
//         }
//         return true;
//     });
// }

// const getSigningKeys = (header, callback) => {
//     var client = jwksClient({
//         jwksUri: DISCOVERY_KEYS_ENDPOINT
//     });

//     client.getSigningKey(header.kid, function (err, key) {
//         var signingKey = key.publicKey || key.rsaPublicKey;
//         console.log('Signing key: ' + signingKey);
//         callback(null, signingKey);
//     });
// }

module.exports = async function (context, req) {
    try {
        context.log.info('Function execution started.');

        // Check if the Authorization header is present
        if (!req.headers.authorization) {
            context.log.error('No access token provided.');
            context.res = {
                status: 401,
                body: 'No access token provided'
            };
            return;
        }

        // Extract bearer token
        const [bearer, token] = req.headers.authorization.split(' ');
        context.log.info(`Token received: ${token}`);

        // Get Graph Token
        context.log.info('Retrieving Graph token...');

        context.res = {
            status: 200,
            headers: {
                'Content-Type': 'application/json'
            },
            body: {
                value: [{displayName :"Came from ListContainers", id: 1}]
            }
        };


        // const [graphSuccess, graphTokenResponse] = await getGraphToken(cca, token);

        // if (!graphSuccess) {
        //     context.log.error('Failed to retrieve Graph token.');
        //     context.res = graphTokenResponse;
        //     return;
        // }

        // // Set up Graph client options
        // context.log.info('Initializing Graph client...');
        // const authProvider = (callback) => {
        //     callback(null, graphTokenResponse);
        // };

        // let options = {
        //     authProvider,
        //     defaultVersion: 'beta'
        // };

        // // Call Microsoft Graph API
        // context.log.info('Calling Microsoft Graph API...');
        // const graph = Graph.Client.init(options);
        // const res = await graph
        //     .api(`storage/fileStorage/containers?$filter=containerTypeId eq ${process.env["APP_CONTAINER_TYPE_ID"]}`)
        //     .get();

        // // Log success
        // context.log.info('Graph API call successful.');

        // // Return successful response
        // context.res = {
        //     body: res
        // };
    } catch (error) {
        // Log the error in detail
        context.log.error('An error occurred:', error);

        // Return the error to the consumer
        context.res = {
            status: 200,
            body: {
                message: `An unexpected error occurred. ${error.message}`,
                details: error.message
            }
        };
    }
};

