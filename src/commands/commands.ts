/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, Office */

import { Client, Attribute, symcrypt } from "@e4a/irmaseal-client"

import * as msal from "@azure/msal-browser"
import { AccountInfo } from "@azure/msal-browser"
import { LogLevel } from "msal"

// eslint-disable-next-line no-undef
var Buffer = require("buffer/").Buffer

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
                // eslint-disable-next-line no-unused-vars
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

var mailboxItem
var globalEvent

Office.initialize = () => {
    Office.onReady(() => {
        mailboxItem = Office.context.mailbox.item
    })
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
// eslint-disable-next-line no-unused-vars
function encrypt(event: Office.AddinCommands.Event) {
    const message: Office.NotificationMessageDetails = {
        type:
            Office.MailboxEnums.ItemNotificationMessageType
                .InformationalMessage,
        message: "Encrypting email with IRMASeal",
        icon: "Icon.80x80",
        persistent: true,
    }

    globalEvent = event

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
                    encryptAndsendMail(accessToken)
                })
                .catch(function (error) {
                    //Acquire token silent failure, send an interactive request
                    if (error instanceof msal.BrowserAuthError) {
                        // show login popup
                        showLoginPopup("/fallbackauthdialog.html")
                    }
                    console.log("[encrypt] Browserautherror: ", error)
                })
        } else {
            const accountId = tokenResponse.account.homeAccountId
            const myAccount: AccountInfo = msalInstance.getAccountByHomeId(
                accountId
            )
            console.log("[encrypt] having token for account id ", myAccount)
        }
    })
}

function getRecipientEmail(): Promise<string> {
    return new Promise(function (resolve, reject) {
        mailboxItem.to.getAsync((recipients) => {
            const recipientEmail: string = recipients.value[0].emailAddress
            if (recipientEmail !== "") resolve(recipientEmail)
            else reject("No recipient email")
        })
    })
}

async function getMailBody(): Promise<string> {
    return new Promise(function (resolve, reject) {
        mailboxItem.body.getAsync(Office.CoercionType.Text, (asyncResult) => {
            const body: string = asyncResult.value
            if (body !== "") resolve(body)
            else reject("No body in email")
        })
    })
}

async function getMailSubject(): Promise<string> {
    return new Promise(function (resolve, reject) {
        mailboxItem.subject.getAsync((asyncResult) => {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                reject("Subject async failed")
            } else {
                const subject: string = asyncResult.value
                if (subject !== "") resolve(subject)
                else reject("No subject in email")
            }
        })
    })
}

async function encryptAndsendMail(_token) {
    const recipientEmail = await getRecipientEmail() //.catch(e => console.error(e)) // mailboxItem.to.getAsync()

    console.log("Recipient: ", recipientEmail)

    const identity: Attribute = {
        type: "pbdf.sidn-pbdf.email.email",
        value: recipientEmail,
    }

    const mailBody = await getMailBody()
    const mailSubject = await getMailSubject()

    console.log("Mailbody: ", mailBody)

    const client = await Client.build("https://irmacrypt.nl/pkg") //.then((client) => {

    const bytes = new TextEncoder().encode(mailBody)

    const meta = client.createMetadata(identity)
    const metadata = meta.metadata.to_json()

    console.log("meta.header: ", meta.header)
    console.log("meta.keys: ", meta.keys)
    console.log("meta.metadata: ", metadata)
    console.log("nonce: ", metadata.iv)

    const ct = await symcrypt(meta.keys, metadata.iv, meta.header, bytes) //then((ct) => {
    console.log("ct :", ct)
    const b64encoded = Buffer.from(ct).toString("base64")

    //const restHost = Office.context.mailbox.restUrl + "/v2.0/me/sendMail"
    const sendMessageUrl = "https://graph.microsoft.com/beta/me/sendMail"

    console.log("Trying to send email via ", sendMessageUrl)

    let message = {
        message: {
            subject: mailSubject,
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

    const BOUNDARY = "--+IRMASEAL+--"
    const encryptedData = b64encoded.replace(/(.{80})/g, "$1\n")

    const headers = {
        Subject: `${mailSubject}`,
        To: `${recipientEmail}`,
        From: `${Office.context.mailbox.userProfile.emailAddress}`,
        "MIME-Version": "1.0",
        "Content-Type": `multipart/encrypted; protocol="application/irmaseal"; boundary=${BOUNDARY}`,
    }

    var headerStr = ""
    for (const [k, v] of Object.entries(headers)) {
        headerStr += `${k}: ${v}\r\n`
    }
    headerStr += "\r\n\r\n"

    console.log("Encrypted data: ", encryptedData)

    var content = headerStr
    content += "Content-Type: text/plain\r\n\r\n"
    content += "This is an IRMAseal/MIME encrypted message.\r\n\r\n"
    content += `--${BOUNDARY}\r\n`
    content += "Content-Type: application/irmaseal\r\n"
    content += "Content-Transfer-Encoding: base64\r\n\r\n"
    content += "Version: 1\r\n\r\n"
    content += `--${BOUNDARY}\r\n`
    content += "Content-Type: application/octet-stream\r\n"
    content += "Content-Transfer-Encoding: base64\r\n\r\n"
    content += `${encryptedData}\r\n\r\n`
    content += `--${BOUNDARY}--\r\n`

    message = Buffer.from(content).toString("base64")

    console.log("Email content: ", message)

    $.ajax({
        type: "POST",
        contentType: "text/plain",
        url: sendMessageUrl,
        data: message,
        headers: {
            Authorization:
                "Bearer " +
                // token,
                "EwB4A8l6BAAU6k7+XVQzkGyMv7VHB/h4cHbJYRAAAapYNkliFsERNAyVY90gq7LRUY8U07UJVKFwC9f+LnG8T/PGBIDogIYqVBQ2QjzRPqvCOUzhAeGEU3+hKaA44lU7LLVBbltcsANMuyacFymeZqerNZC2buuPkKuiAfU/0qAh+KwPMdtvCP2rH3rUD28JsBa5knnwPNIMVUh78ROwnc2MVgeqOwEwxPAYBP9T8Q3GC73mOvh7ESc7Kzvibfr4PbxdCZhJHfb5Ur3l5ZnWzbCuZdo/+KNDW9ln8NbQ30I4OKOhKi1vue2nHFGtQEK2Zqip1l2YHLaM536a8p3ENd2adZG59hH11MrS0KyMPL8/on5E2xRN1iwvl9wbAk4DZgAACHpN3BGkydabSALTgbsJeaJQ4dZUD/hLKO65NM3fkIuFviR7XnmGIA2sBGnPsb1Ge1QFItdvA1cpYJjG7Em35fPHRkZMjDWcJ50w3u6r21t2JkBI1am9MO/hqKY/GaJtxs6lh7yw2yzDGcfVmty1custdLA9G52Y6xSix7o9nivLo9NE//+gafnguQRQy5RiIxcRNm9QQKdc64L419etlAlBCsOw+07drvJm+9k7ADFWyeH8j6H5Nv1jdfDgBb+E+wI+DV//jmDL7zd9Z8VuNtnnXLZwl/60aEeMxcL7xlqafKydXkOMi3sIumbvwaE+vQ2fCzd8hVsh4q2s3aoGabZWu4MuI7pLXKICcEyFxXhOZKaN/OITfkHNC4yzwXTY2liclT/W2PO4DTW4l4FfvsMmYfjmxX3bVxXuc0niHuyt5zLJ1zb7kPM6hIL3grLZC5QWvai+M05hMAqhwzTzF+9/DBudVOcVSXTsJz/cm37RP0C7d/XT3ykVG2BbNWRAdcbhylZ19s9J7h/rVGjpWoZRWTeE+v92DGkuNLMalQdPN1M/IpP3JA3s316icdnAhkLyRs5aLtgy8wMlzD0O8m6vGYZyPYmMOYiE3BvbUHRS86Tly2Le1Mk7/3ht6bR6FzLry7VYYb0Y5t+LGw59O5Z1kVALl3+JdZDZq7hxzfhEG6dWi/68KRvA0nMDcA0CNO7Lwnf1182DGBwADIrgcyfB/0b81omH4J2Qa2dpWneIMrofpatwVbSIGQPv2KkkHjVsJ/k1JjSTx49ZBYxMed8DSI8C",
        },
        success: function (success) {
            console.log("Sendmail success: ", success)

            const successMsg: Office.NotificationMessageDetails = {
                type:
                    Office.MailboxEnums.ItemNotificationMessageType
                        .InformationalMessage,
                message: "Successfully encrypted and send email",
                icon: "Icon.80x80",
                persistent: true,
            }

            Office.context.mailbox.item.notificationMessages.replaceAsync(
                "action",
                successMsg
            )

            globalEvent.completed()
        },
    }).fail(function ($xhr) {
        var data = $xhr.responseJSON
        console.log("Ajax error: ", data)
        setEventError()
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
