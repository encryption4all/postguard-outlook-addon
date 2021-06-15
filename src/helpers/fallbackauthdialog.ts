/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to use MSAL.js to get an access token to Microsoft Graph an pass it to the task pane.
 */

/* global console, localStorage, Office */

import * as msal from "@azure/msal-browser"
import { LogLevel } from "@azure/msal-browser"

const msalConfig = {
    auth: {
        clientId: "6ee2a054-1d61-405d-8e5d-c2daf25c5833",
        authority:
            "https://login.microsoftonline.com/f7a5af59-797f-4868-9c79-0c51006c58f6",
    },
    cache: {
        cacheLocation: "sessionStorage",
    },
    system: {
        loggerOptions: {
            loggerCallback: (
                level: LogLevel,
                message: string,
                containsPii: boolean
            ): void => {
                switch (level) {
                    case LogLevel.Error:
                        console.error("[logger]", message)
                        return
                    case LogLevel.Info:
                        console.info("[logger]", message)
                        return
                    case LogLevel.Verbose:
                        console.debug("[logger]", message)
                        return
                    case LogLevel.Warning:
                        console.warn("[logger]", message)
                        return
                }
            },
            piiLoggingEnabled: false,
        },
    },
}

const requestObj = {
    scopes: ["user.read", "mail.send"],
}

const msalInstance = new msal.PublicClientApplication(msalConfig)

Office.initialize = function () {
    if (Office.context.ui.messageParent) {
        msalInstance.handleRedirectPromise().then((response) => {
            // The very first time the add-in runs on a developer's computer, msal.js hasn't yet
            // stored login data in localStorage. So a direct call of acquireTokenRedirect
            // causes the error "User login is required". Once the user is logged in successfully
            // the first time, msal data in localStorage will prevent this error from ever hap-
            // pening again; but the error must be blocked here, so that the user can login
            // successfully the first time. To do that, call loginRedirect first instead of
            // acquireTokenRedirect.
            if (
                response !== null ||
                localStorage.getItem("loggedIn") === "yes"
            ) {
                authCallback(null, response)
                msalInstance.acquireTokenRedirect(requestObj)
            } else {
                // This will login the user and then the (response.tokenType === "id_token")
                // path in authCallback below will run, which sets localStorage.loggedIn to "yes"
                // and then the dialog is redirected back to this script, so the
                // acquireTokenRedirect above runs.
                msalInstance.loginRedirect(requestObj)
            }
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
            console.log(response.idToken.rawIdToken)
            localStorage.setItem("loggedIn", "yes")
        } else {
            console.log("token type is:" + response.tokenType)
            Office.context.ui.messageParent(
                JSON.stringify({
                    status: "success",
                    result: response.accessToken,
                })
            )
        }
    }
}
