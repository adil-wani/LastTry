import { ConfidentialClientApplication, LogLevel } from '@azure/msal-node';
import jwt from 'jsonwebtoken';
import jwksClient from 'jwks-rsa';
import 'isomorphic-fetch';
import { Client as GraphClient } from '@microsoft/microsoft-graph-client';
import { getGraphToken } from '../utils/auth.js';

const DISCOVERY_KEYS_ENDPOINT = process.env["DISCOVERY_KEYS_ENDPOINT"];
const config = {
    auth: {
        clientId: process.env["APP_CLIENT_ID"],
        authority: process.env["APP_AUTHORITY"],
        audience: process.env["APP_AUDIENCE"],
        clientSecret: process.env["APP_CLIENT_SECRET"]
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: LogLevel.Verbose,
        }
    }
};
const cca = new ConfidentialClientApplication(config);

const isJwtValid = (token) => {
    if (!token) {
        return false;
    }
    const validationOptions = {
        algorithms: ['RS256'],
        audience: config.auth.audience,
        issuer: config.auth.issuer
    };
    jwt.verify(token, getSigningKeys, validationOptions, (err, payload) => {
        if (err) {
            console.log(err);
            return false;
        }
        return true;
    });
};

const getSigningKeys = (header, callback) => {
    const client = jwksClient({
        jwksUri: DISCOVERY_KEYS_ENDPOINT
    });

    client.getSigningKey(header.kid, (err, key) => {
        const signingKey = key.publicKey || key.rsaPublicKey;
        console.log('Signing key: ' + signingKey);
        callback(null, signingKey);
    });
};

export async function run(context, req) {
    try {
        context.log.info('Function execution started.');

        if (!req.headers.authorization) {
            context.log.error('No access token provided.');
            context.res = {
                status: 401,
                body: 'No access token provided'
            };
            return;
        }

        const [bearer, token] = req.headers.authorization.split(' ');
        context.log.info(`Token received: ${token}`);

        context.log.info('Retrieving Graph token...');

        const [graphSuccess, graphTokenResponse] = await getGraphToken(cca, token);

        if (!graphSuccess) {
            context.log.error('Failed to retrieve Graph token.');
            context.res = graphTokenResponse;
            return;
        }

        context.log.info('Initializing Graph client...');
        const authProvider = (callback) => {
            callback(null, graphTokenResponse);
        };

        const options = {
            authProvider,
            defaultVersion: 'beta'
        };

        context.log.info('Calling Microsoft Graph API...');
        const graph = GraphClient.init(options);
        const res = await graph
            .api(`storage/fileStorage/containers?$filter=containerTypeId eq ${process.env["APP_CONTAINER_TYPE_ID"]}`)
            .get();

        context.log.info('Graph API call successful.');

        context.res = {
            body: res
        };
    } catch (error) {
        context.log.error('An error occurred:', error);

        context.res = {
            status: 200,
            body: {
                message: `An unexpected error occurred. ${error.message}`,
                details: error.message
            }
        };
    }
}
