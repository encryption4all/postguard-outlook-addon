/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, Office */

import { Client, Attribute } from "@e4a/irmaseal-client"
import { ComposeMail } from "@e4a/irmaseal-mail-utils"
import { CryptifyApiWrapper } from "@e4a/cryptify-api-wrapper/dist/cryptify-api-wrapper"
import { merge } from "jquery"

// eslint-disable-next-line no-undef
var Buffer = require("buffer/").Buffer

var loginDialog
var mailboxItem
var globalEvent

// in bytes (1024 x 1024 = 1 MB)
const MAX_ATTACHMENT_SIZE = 1024 * 1024

Office.initialize = () => {
    Office.onReady(() => {
        mailboxItem = Office.context.mailbox.item

        delete window.alert // assures alert works
        delete window.confirm // assures confirm works
        delete window.prompt // assures prompt works
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
        mailboxItem.body.getAsync(Office.CoercionType.Html, (asyncResult) => {
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

interface IAttachmentContent {
    filename: string
    content: string
}

async function getMailAttachments(): Promise<IAttachmentContent[]> {
    return new Promise(function (resolve, reject) {
        mailboxItem.getAttachmentsAsync(async (asyncResult) => {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                reject("Attachments async failed")
            } else {
                if (asyncResult.value.length > 0) {
                    let attachmentsArray = []
                    let content = ""
                    for (var i = 0; i < asyncResult.value.length; i++) {
                        var attachment = asyncResult.value[i]
                        content = await getMailAttachmentContent(attachment.id)
                        attachmentsArray.push({
                            filename: attachment.name,
                            content: content,
                        })
                    }
                    resolve(attachmentsArray)
                } else {
                    reject("No attachments in email")
                }
            }
        })
    })
}

async function getMailAttachmentContent(attachmentId: string): Promise<string> {
    return new Promise(function (resolve, reject) {
        mailboxItem.getAttachmentContentAsync(attachmentId, (asyncResult) => {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                reject("Attachment content async failed")
            } else {
                if (asyncResult.value.content.length > 0) {
                    resolve(asyncResult.value.content)
                } else
                    reject(
                        "No attachment content in attachment with id" +
                            attachmentId
                    )
            }
        })
    })
}

// Encrypts and sends the mail
async function encryptAndsendMail(token) {
    const recipientEmail = await getRecipientEmail() //.catch(e => console.error(e)) // mailboxItem.to.getAsync()
    const sender = Office.context.mailbox.userProfile.emailAddress

    console.log("Recipient: ", recipientEmail)

    const identity: Attribute = {
        type: "pbdf.sidn-pbdf.email.email",
        value: recipientEmail,
    }

    let mailBody = await getMailBody()
    // extract HTML within <body>
    const pattern = /<body[^>]*>((.|[\n\r])*)<\/body>/im
    const arrayMatches = pattern.exec(mailBody)
    mailBody = arrayMatches[1]

    const mailSubject = await getMailSubject()

    let attachments: IAttachmentContent[]
    await getMailAttachments()
        .then((attas) => (attachments = attas))
        .catch((error) => console.log(error))

    console.log("Mail subject: ", mailSubject)

    const client = await Client.build("https://irmacrypt.nl/pkg")
    /*const controller = new AbortController()
    const cryptifyApiWrapper = new CryptifyApiWrapper(
        client,
        recipientEmail,
        sender,
        "https://dellxps"
    )*/

    const meta = client.createMetadata(identity)
    const metadata = meta.metadata.to_json()

    const keys = meta.keys
    const header: Uint8Array = meta.header
    const iv = metadata.iv

    console.log("meta.header: ", meta.header)
    console.log("meta.keys: ", meta.keys)
    console.log("meta.metadata: ", metadata)
    console.log("nonce: ", metadata.iv)

    const composeMail = new ComposeMail()
    composeMail.addRecipient(recipientEmail)
    composeMail.setVersion("1")
    composeMail.setSubject(mailSubject)
    composeMail.setSender(sender)

    if (attachments !== undefined) {
        for (let i = 0; i < attachments.length; i++) {
            const attachment = attachments[i]

            let useCryptify = false
            const fileBlob = new Blob([attachment.content], {
                type: "application/octet-stream",
            })
            const file = new File([fileBlob], attachment.filename, {
                type: "application/octet-stream",
            })

            // if attachment is too large, ask user if it should be encrypted via Cryptify
            /*
            if (fileBlob.size > MAX_ATTACHMENT_SIZE) {
                // TODO: Add confirmation dialog (https://theofficecontext.com/2017/06/14/dialogs-in-officejs/)
                console.log(
                    `Attachment ${attachment.filename} larger than 1 MB`
                )
                useCryptify = true
                const downloadUrl = await cryptifyApiWrapper.encryptAndUploadFile(
                    file,
                    controller
                )
                mailBody += `<p><a href="${downloadUrl}">Download ${attachment.filename} via Cryptify</a></p>`
            }
            */

            if (!useCryptify) {
                let nonce = new Uint8Array(8)
                nonce = window.crypto.getRandomValues(nonce)

                const counter = new Uint8Array(8)
                const mergedArray = new Uint8Array([...nonce, ...counter])

                const input = new TextEncoder().encode(attachment.content)

                console.log("Attachment bytes length: ", input.byteLength)

                const attachmentCT = await client.symcrypt({
                    keys,
                    iv: mergedArray,
                    header: mergedArray,
                    input,
                })
                composeMail.addAttachment(attachmentCT, attachment.filename)
            }
        }
    }

    console.log("Mailbody: ", mailBody)

    const input = new TextEncoder().encode(mailBody)
    console.log("Body bytes length: ", input.byteLength)
    const ct = await client.symcrypt({ keys, iv, header, input })
    composeMail.setCiphertext(ct)

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
