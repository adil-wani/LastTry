const msal = require('@azure/msal-node');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');
require('isomorphic-fetch'); // required for graph library and SharePoint calls
const Graph = require('@microsoft/microsoft-graph-client');
//const { getGraphToken } = require('../utils/auth');

const DISCOVERY_KEYS_ENDPOINT = process.env.APP_CLIENT_ID;

console.log('DISCOVERY_KEYS_ENDPOINT:', DISCOVERY_KEYS_ENDPOINT);
const config = {
    auth: {
        clientId: process.env["APP_CLIENT_ID"],
        authority: process.env["APP_AUTHORITY"],
        audience: process.env["APP_AUDIENCE"],
        clientSecret: process.env["APP_CLIENT_SECRET"],
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Verbose,
        }
    }
};

const cca = new msal.ConfidentialClientApplication(config);

const getSigningKeys = (header, callback) => {
    const client = jwksClient({
        jwksUri: DISCOVERY_KEYS_ENDPOINT,
    });

    client.getSigningKey(header.kid, (err, key) => {
        const signingKey = key.publicKey || key.rsaPublicKey;
        console.log('Signing key:', signingKey);
        callback(null, signingKey);
    });
};

const isJwtValid = (token) => {
    if (!token) return false;

    const validationOptions = {
        algorithms: ['RS256'],
        audience: config.auth.audience,
        issuer: config.auth.authority,
    };

    try {
        jwt.verify(token, getSigningKeys, validationOptions);
        return true;
    } catch (error) {
        console.error('JWT validation error:', error);
        return false;
    }
};

const { app } = require('@azure/functions');


async function getGraphToken(cca, token) {
    try {
        const graphTokenRequest = {
            oboAssertion: token,
            scopes: ["Sites.Read.All", "FileStorageContainer.Selected"]
        };
        const graphToken = (await cca.acquireTokenOnBehalfOf(graphTokenRequest)).accessToken;
        return [true, graphToken];
    } catch (error) {
        const errorResult = {
            status: 500,
            body: JSON.stringify({
                message: 'Unable to generate graph obo token: ' + error.message,
                providedToken: token
            })
        };
        return [false, errorResult];
    }
}



app.http('ListContainers', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log(`HTTP function processed request for URL: "${request.url}"`);

        // Check for authorization header
        const authHeader = request.headers.get('authorization');
        if (!authHeader) {
            return {
                status: 401,
                body: 'No access token provided.',
            };
        }

        const token = authHeader.split(' ')[1];
        context.log(`Token received: ${token}`);

        // Validate JWT
        // if (!isJwtValid(token)) {
        //     return {
        //         status: 401,
        //         body: 'Invalid token.',
        //     };
        // }

        // Retrieve Graph token
        context.log('Retrieving Graph token...');

        try {
        const [graphSuccess, graphTokenResponse] = await getGraphToken(cca, token);

        if (!graphSuccess) {
            context.log('Failed to retrieve Graph token.');
            return graphTokenResponse;
        }

        // Set up Graph client
        context.log('Initializing Graph client...');
        const authProvider = (callback) => {
            callback(null, graphTokenResponse);
        };

        const options = {
            authProvider,
            defaultVersion: 'beta',
        };

        // Call Microsoft Graph API
        context.log('Calling Microsoft Graph API...');
        const graphClient = Graph.Client.init(options);
        const res = await graphClient
            .api(`storage/fileStorage/containers?$filter=containerTypeId eq ${process.env["APP_CONTAINER_TYPE_ID"]}`)
            .get();

        context.log('Graph API call successful.');

        return {
            status: 200,
            body: JSON.stringify(res),
        };
    } catch (error) {

        return {
            status: 500,
            body: error.message,
        };
    }
    }
});
