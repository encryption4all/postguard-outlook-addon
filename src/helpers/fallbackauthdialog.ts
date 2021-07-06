/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to use MSAL.js to get an access token to Microsoft Graph an pass it to the task pane.
 */

/* global Office */

import * as msal from "@azure/msal-browser"
import { LogLevel, SilentRequest } from "@azure/msal-browser"

let logginger: string

const msalConfig = {
    auth: {
        clientId: "6ee2a054-1d61-405d-8e5d-c2daf25c5833",
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true,
    },
    system: {
        loggerOptions: {
            loggerCallback: (
                level: LogLevel,
                message: string,
                // eslint-disable-next-line no-unused-vars
                containsPii: boolean
            ): void => {
                switch (level) {
                    case LogLevel.Error:
                        console.error("[fallback]", message)
                        return
                    case LogLevel.Info:
                        console.info("[fallback]", message)
                        return
                    case LogLevel.Verbose:
                        console.debug("[fallback]", message)
                        return
                    case LogLevel.Warning:
                        console.warn("[fallback]", message)
                        return
                }
            },
            piiLoggingEnabled: false,
        },
    },
}

const requestObj = {
    scopes: ["Mail.Send", "openid", "profile", "offline_access"],
}

const msalInstance = new msal.PublicClientApplication(msalConfig)

Office.initialize = function () {
    if (Office.context.ui.messageParent) {
        msalInstance.handleRedirectPromise().then((response) => {
            const silentFlowRequest: SilentRequest = {
                scopes: requestObj.scopes,
                account: msalInstance.getAllAccounts()[0],
                forceRefresh: false,
            }

            msalInstance
                .acquireTokenSilent(silentFlowRequest)
                .then(function (accessTokenResponse) {
                    logginger = "Silent!"
                    authCallback(null, accessTokenResponse)
                })
                // eslint-disable-next-line no-unused-vars
                .catch(function (error) {
                    let accountObj
                    if (response) {
                        accountObj = response.account
                    } else {
                        accountObj = msalInstance.getAllAccounts()[0]
                    }

                    if (accountObj && response) {
                        logginger = "account and response"
                        authCallback(null, response)
                    }
                    // The very first time the add-in runs on a developer's computer, msal.js hasn't yet
                    // stored login data in localStorage. So a direct call of acquireTokenRedirect
                    // causes the error "User login is required". Once the user is logged in successfully
                    // the first time, msal data in localStorage will prevent this error from ever hap-
                    // pening again; but the error must be blocked here, so that the user can login
                    // successfully the first time. To do that, call loginRedirect first instead of
                    // acquireTokenRedirect.
                    else if (localStorage.getItem("loggedIn") === "yes") {
                        msalInstance.acquireTokenRedirect(requestObj)
                        logginger = "loggedin yes"
                    } else {
                        // This will login the user and then the (response.tokenType === "id_token")
                        // path in authCallback below will run, which sets localStorage.loggedIn to "yes"
                        // and then the dialog is redirected back to this script, so the
                        // acquireTokenRedirect above runs.
                        msalInstance.loginRedirect(requestObj)
                        logginger = "loggedin no"
                    }
                })
        })
    }
}

function authCallback(error, response) {
    if (error) {
        console.log(error)
        Office.context.ui.messageParent(
            JSON.stringify({ status: "failure", result: error })
        )
    } else {
        if (response.tokenType === "id_token") {
            localStorage.setItem("loggedIn", "yes")
            console.log("id_token!")
            console.log(response.idToken.rawIdToken)
        } else {
            Office.context.ui.messageParent(
                JSON.stringify(
                    {
                        status: "success",
                        result: response,
                        logging: logginger,
                    },
                    replacer
                )
            )
        }
    }
}
