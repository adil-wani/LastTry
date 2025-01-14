import { ConfidentialClientApplication } from '@azure/msal-node';
import jwt from 'jsonwebtoken';
import jwksClient from 'jwks-rsa';
import 'isomorphic-fetch';
import { Client as GraphClient } from '@microsoft/microsoft-graph-client';

export async function getGraphToken(cca, token) {
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
