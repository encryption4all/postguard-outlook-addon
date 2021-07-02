/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, Office */

import { Client, Attribute, symcrypt } from "@e4a/irmaseal-client"

import * as msal from "@azure/msal-browser"
import { AccountInfo } from "@azure/msal-browser"
import { LogLevel } from "msal"
import { ComposeMail } from "@e4a/irmaseal-mail-utils"

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

async function encryptAndsendMail(token) {
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

    const composeMail = new ComposeMail()
    composeMail.addRecipient(recipientEmail)
    composeMail.setVersion("1")
    composeMail.setSubject(mailSubject)
    composeMail.setCiphertext(ct)
    composeMail.setSender(Office.context.mailbox.userProfile.emailAddress)

    const message = Buffer.from(composeMail.getMimeMail()).toString("base64")

    //const restHost = Office.context.mailbox.restUrl + "/v2.0/me/sendMail"
    const sendMessageUrl = "https://graph.microsoft.com/beta/me/sendMail"

    console.log("Trying to send email via ", sendMessageUrl)

    $.ajax({
        type: "POST",
        contentType: "text/plain",
        url: sendMessageUrl,
        data: message,
        headers: {
            Authorization:
                "Bearer " +
                // token,
                "EwB4A8l6BAAU6k7+XVQzkGyMv7VHB/h4cHbJYRAAASn4oCFTDlgDDBWS1TVQCZoprMNwOWLZyGWyFZQFAjGvl0FCiXVbnSRXCR7daUJVN3R3bl29Qrhw40HkcI3N1AcgpqsKfh74MmtP91xDt0TF+QUZlOMkEJSHEnZFYn7bBcC2X0eQ0OEXLOTmBtIkcdzspwYW6c4gO+PJimH6F8ZjW0PXE7X6Vkr/LCdQ79ighBhhu5QSl98YieV3C7xKwgD7t+KyANcZ2k0cJajS4QnL3g45M3SbO0+n8/PAQn9AxvgJtLHecq1STx1yhOvz45L9Wo0r5QFJJjuB6nV5uKCUssufHEjmHphbKL8YfBr/QhUffr3KVomluenCZXoorlsDZgAACOLPBNiZk8pUSAJIKPB5NJhT/BfkexxGTRBe5rZOYEUvVL13eUbCDNDmHb71EMUmbPZnD8alT4wYpsHlRX2cZ4EJQpGYzu08hd5egJXF9adJtiW/W1F15bODOPw74XNRGk65B8MXVfx7ZUPIdatuQseKaaAqpUV77qDE8MDkgKeAbS+vItOHOhrt7pWX87aZYc/yvU1uW1j+706KLDKEXT+hbxd5ZFwEo8VAOI4DTeGVNVU5APK3Gr+poTCnK6LJTxmmoBuFW8NnTANN/YSmxcKhw+0KfO2JpxPlHMiDEoMwc/LB+EzfGdl/N5X/m0QtFtbFpf0VoSBcj3REQDdFU25TUwCe/VA5+SRz5tgMzT9tgpKTUUUrwzC+1KgnaHX5ZblEgiMzxt3ymQRhijzOrPSwCjssrKZjgCS1SOslywxDptgxuu1Cf/mowP7F6Zg44pJGvnyCH8mAg1r0oo8XxTd7nWujujxVn3eVQSjE6+nKUkT1X8WvZxdiZSvihXGR8ZTv2TgcGGU45VuFcX9XflYGNwMCSd897cPeR4TYo1WT5C8ZZ9h/oA5DBTlXRrcC70KzIzND6EGs2rxo3rcEs7zJouMualsqWf4SPx66Z0wDg4IrDpR73eYydoXRt+J7ClQVIE9TCw97KA3ZXihpe1HhfwoYEgFag1uE2SpSDQsXhADGYDJp5tzzv5PHF5GuHImrzmDGJsa0c1aDtKnaFPiZX1gSH0W0aVxvFXuTXWfKMHzeH5Y0AuvNRRng8glTdmyJfT4kbeqbeMG4bcezyGaY0I8C",
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
