/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { Client, Attribute, symcrypt } from "@e4a/irmaseal-client"

import * as msal from "@azure/msal-browser"
import { AccountInfo } from "@azure/msal-browser"
import { LogLevel } from "msal"

var Buffer = require("buffer/").Buffer
var fs = require("browserify-fs")

const filename = "encrypt.log"

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

var loginDialog

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
async function processMessage(arg) {
    console.log("Message received in processMessage: " + JSON.stringify(arg))
    let messageFromDialog = JSON.parse(arg.message)

    if (messageFromDialog.status === "success") {
        // We now have a valid access token.
        loginDialog.close()
        console.log("Valid token: ", JSON.stringify(messageFromDialog.result))
        encryptAndsendMail(messageFromDialog.result)
    } else {
        // Something went wrong with authentication or the authorization of the web application.
        loginDialog.close()
        console.log(
            "Message from dialog error: ",
            JSON.stringify(messageFromDialog.error.toString())
        )
    }
}

// Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
function showLoginPopup(url) {
    var fullUrl =
        location.protocol +
        "//" +
        location.hostname +
        (location.port ? ":" + location.port : "") +
        url

    // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
    Office.context.ui.displayDialogAsync(
        fullUrl,
        { height: 60, width: 30 },
        function (result) {
            console.log("Dialog has initialized. Wiring up events")
            loginDialog = result.value
            loginDialog.addEventHandler(
                Office.EventType.DialogMessageReceived,
                processMessage
            )
        }
    )
}

Office.initialize = () => {
    Office.onReady(() => {})
}

function setEventError() {
    const message: Office.NotificationMessageDetails = {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: "Error during encryption, please contact your administrator.",
    }

    Office.context.mailbox.item.notificationMessages.replaceAsync(
        "action",
        message
    )
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function encrypt(event: Office.AddinCommands.Event) {
    const message: Office.NotificationMessageDetails = {
        type:
            Office.MailboxEnums.ItemNotificationMessageType
                .InformationalMessage,
        message: "Encrypting email with IRMASeal",
        icon: "Icon.80x80",
        persistent: true,
    }

    Office.context.mailbox.item.notificationMessages.replaceAsync(
        "action",
        message
    )

    const msalInstance = new msal.PublicClientApplication(msalConfig)
    const accessTokenRequest = {
        scopes: ["user.read", "mail.send"],
    }

    msalInstance.handleRedirectPromise().then((tokenResponse) => {
        if (tokenResponse === null) {
            msalInstance
                .acquireTokenSilent(accessTokenRequest)
                .then(function (accessTokenResponse) {
                    // Acquire token silent success
                    let accessToken = accessTokenResponse.accessToken
                    // Call your API with token
                    console.log("[encrypt] pca token: ", accessToken)
                    //encryptAndsendMail(accessToken)
                })
                .catch(function (error) {
                    //Acquire token silent failure, send an interactive request
                    if (error instanceof msal.BrowserAuthError) {
                        showLoginPopup("/fallbackauthdialog.html")
                        /*msalInstance
                            .acquireTokenRedirect(accessTokenRequest)
                            .catch(function (error) {
                                // Acquire token interactive failure
                                console.log(error)
                            })*/
                    }
                    console.log("[encrypt] Browserautherror: ", error)
                })
        } else {
            const accountId = tokenResponse.account.homeAccountId
            const myAccount: AccountInfo = msalInstance.getAccountByHomeId(
                accountId
            )
            console.log("[encrypt] having token for accid ", myAccount)
        }
    })
}

function encryptAndsendMail(token) {
    const mailboxItem = Office.context.mailbox.item

    mailboxItem.to.getAsync((recipients) => {
        const recipientEmail = recipients.value[0].emailAddress

        console.log("Recipient: ", recipientEmail)

        const identity: Attribute = {
            type: "pbdf.sidn-pbdf.email.email",
            value: recipientEmail,
        }

        mailboxItem.body.getAsync(Office.CoercionType.Text, (asyncResult) => {
            console.log("Mailbody: ", asyncResult.value)

            Client.build("https://irmacrypt.nl/pkg").then((client) => {
                const bytes = new TextEncoder().encode(asyncResult.value)

                const meta = client.createMetadata(identity)
                const metadata = meta.metadata.to_json()

                console.log("meta.header: ", meta.header)
                console.log("meta.keys: ", meta.keys)
                console.log("meta.metadata: ", metadata)
                console.log("nonce: ", metadata.iv)

                symcrypt(meta.keys, metadata.iv, meta.header, bytes).then(
                    (ct) => {
                        console.log("ct :", ct)
                        const b64encoded = Buffer.from(ct).toString("base64")

                        //const restHost = Office.context.mailbox.restUrl
                        const sendMessageUrl =
                            "https://graph.microsoft.com/v1.0/me/sendMail"
                        //  restHost + "/v2.0/me/sendMail"

                        console.log("Trying to send email via ", sendMessageUrl)

                        let message = {
                            message: {
                                subject: "test", //Office.context.mailbox.item.subject,
                                body: {
                                    contentType: "Text",
                                    content: b64encoded,
                                },
                                toRecipients: [
                                    {
                                        emailAddress: {
                                            address: recipientEmail,
                                        },
                                    },
                                ],
                            },
                        }

                        const BOUNDARY = "foo"
                        const encryptedData = b64encoded.replace(
                            /(.{80})/g,
                            "$1\n"
                        )

                        var content =
                            "This is an IRMAseal/MIME encrypted message.\r\n\r\n"
                        content += `--${BOUNDARY}\r\n`
                        content += "Content-Type: application/irmaseal\r\n\r\n"
                        content += "Version: 1\r\n\r\n"
                        content += `--${BOUNDARY}\r\n`
                        content +=
                            "Content-Type: application/octet-stream\r\n\r\n"
                        content += `${encryptedData}\r\n\r\n`
                        content += `--${BOUNDARY}--\r\n`

                        //message = Buffer.from(content).toString("base64")

                        console.log("Email content: ", message)

                        $.ajax({
                            type: "POST",
                            //contentType: "text/plain",
                            contentType: "application/json; charset=utf-8",
                            processData: false,
                            url: sendMessageUrl,
                            data: message,
                            headers: {
                                Authorization: "Bearer " + token,
                                ContentType: "application/json",
                            },
                            success: function (success) {
                                console.log("Sendmail success: ", success)

                                const successMsg: Office.NotificationMessageDetails = {
                                    type:
                                        Office.MailboxEnums
                                            .ItemNotificationMessageType
                                            .InformationalMessage,
                                    message:
                                        "Successfully encrypted and send email",
                                    icon: "Icon.80x80",
                                    persistent: true,
                                }

                                Office.context.mailbox.item.notificationMessages.replaceAsync(
                                    "action",
                                    successMsg
                                )
                            },
                        }).fail(function ($xhr) {
                            var data = $xhr.responseJSON
                            console.log("Ajax error: ", data)
                            setEventError()
                        })
                    }
                )
            })
        })
    })
}

/*
function graphAPITokenCallback() {
    
    const msalInstance = new msal.PublicClientApplication(msalConfig)
    const username = Office.context.mailbox.userProfile.emailAddress // "daniel.ostkamp@outlook.com";
    const myAccount: AccountInfo = msalInstance.getAccountByUsername(username)
    msalInstance.setActiveAccount(myAccount)

    msalInstance.handleRedirectPromise().then((tokenResponse) => {
        let accountObj
        if (tokenResponse) {
            accountObj = tokenResponse.account
        } else {
            accountObj = msalInstance.getAllAccounts()[0]
        }

        if (accountObj && tokenResponse) {
            console.log(
                "[AuthService.init] Got valid accountObj and tokenResponse"
            )
            sendMail(tokenResponse)
        } else if (accountObj) {
            console.log("[AuthService.init] User has logged in, but no tokens.")
            try {
                msalInstance
                    .acquireTokenSilent({
                        account: msalInstance.getAllAccounts()[0],
                        scopes: ["Mail.send"],
                    })
                    .then((tokenResponse) => {
                        sendMail(tokenResponse)
                    })
            } catch (err) {
                console.log("[AuthService.init] Error")
                //await this.app.acquireTokenRedirect({scopes: ["mail.send"]});
            }
        } else {
            console.log(
                "[AuthService.init] No accountObject or tokenResponse present. User must now login."
            )
            msalInstance
                .loginRedirect({ scopes: ["Mail.send"] })
                .then((response) => {
                    console.log("Response: ", response)
                    //sendMail(response)
                })
        }
    })
   }*/

function getGlobal() {
    return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
        ? window
        : typeof global !== "undefined"
        ? global
        : undefined
}

const g = getGlobal() as any

// the add-in command functions need to be available in global scope
g.encrypt = encrypt
