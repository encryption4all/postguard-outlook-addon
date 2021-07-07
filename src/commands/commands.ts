/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, Office */

import { Client, Attribute, symcrypt } from "@e4a/irmaseal-client"
import { ComposeMail } from "@e4a/irmaseal-mail-utils"

// eslint-disable-next-line no-undef
var Buffer = require("buffer/").Buffer

var loginDialog
var mailboxItem
var globalEvent

Office.initialize = () => {
    Office.onReady(() => {
        mailboxItem = Office.context.mailbox.item
    })
}

/**
 * Entry point function.
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

    showLoginPopup("/fallbackauthdialog.html")
}

// Get recipient mail
function getRecipientEmail(): Promise<string> {
    return new Promise(function (resolve, reject) {
        mailboxItem.to.getAsync((recipients) => {
            const recipientEmail: string = recipients.value[0].emailAddress
            if (recipientEmail !== "") resolve(recipientEmail)
            else reject("No recipient email")
        })
    })
}

// Gets the mail body
async function getMailBody(): Promise<string> {
    return new Promise(function (resolve, reject) {
        mailboxItem.body.getAsync(Office.CoercionType.Text, (asyncResult) => {
            const body: string = asyncResult.value
            if (body !== "") resolve(body)
            else reject("No body in email")
        })
    })
}

// Gets the mail subject
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

// Encrypts and sends the mail
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

    const client = await Client.build("https://irmacrypt.nl/pkg")

    const bytes = new TextEncoder().encode(mailBody)

    const meta = client.createMetadata(identity)
    const metadata = meta.metadata.to_json()

    console.log("meta.header: ", meta.header)
    console.log("meta.keys: ", meta.keys)
    console.log("meta.metadata: ", metadata)
    console.log("nonce: ", metadata.iv)

    const ct = await symcrypt(meta.keys, metadata.iv, meta.header, bytes)
    console.log("ct :", ct)

    const composeMail = new ComposeMail()
    composeMail.addRecipient(recipientEmail)
    composeMail.setVersion("1")
    composeMail.setSubject(mailSubject)
    composeMail.setCiphertext(ct)
    composeMail.setSender(Office.context.mailbox.userProfile.emailAddress)

    const message = Buffer.from(composeMail.getMimeMail()).toString("base64")

    const sendMessageUrl = "https://graph.microsoft.com/v1.0/me/sendMail"

    console.log("Trying to send email via ", sendMessageUrl)

    $.ajax({
        type: "POST",
        contentType: "text/plain",
        url: sendMessageUrl,
        data: message,
        headers: {
            Authorization: "Bearer " + token,
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

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
async function processMessage(arg) {
    let messageFromDialog = JSON.parse(arg.message)

    if (messageFromDialog.status === "success") {
        // We now have a valid access token.
        loginDialog.close()
        console.log("Valid token: ", JSON.stringify(messageFromDialog.result))
        console.log("Logginger: ", JSON.stringify(messageFromDialog.logging))
        encryptAndsendMail(messageFromDialog.result.accessToken)
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

function getGlobal() {
    return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
        ? window
        : typeof global !== "undefined"
        ? // eslint-disable-next-line no-undef
          global
        : undefined
}

const g = getGlobal() as any

// the add-in command functions need to be available in global scope
g.encrypt = encrypt
