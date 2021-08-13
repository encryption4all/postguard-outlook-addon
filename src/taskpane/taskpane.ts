/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png"
import "../../assets/icon-32.png"
import "../../assets/icon-80.png"

import "web-streams-polyfill"

import { ReadMail } from "@e4a/irmaseal-mail-utils"

import { Client } from "@e4a/irmaseal-client"

import * as IrmaCore from "@privacybydesign/irma-core"
import * as IrmaClient from "@privacybydesign/irma-client"
import * as IrmaPopup from "@privacybydesign/irma-popup"

/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none"
        document.getElementById("app-body").hidden = false
        document.getElementById("run").onclick = run
    }
})

var item
var mailbox

export async function run() {
    console.log("Run method")

    $(function () {
        getGraphAPIToken()
    })
}

Office.initialize = function () {
    console.log("Initialize")

    item = Office.context.mailbox.item
    mailbox = Office.context.mailbox

    $(function () {
        setItemBody()
    })
}

// Get the body type of the composed item, and set data in
// in the appropriate data type in the item body.
function setItemBody() {
    item.body.getAsync("text", (result) => {
        console.log(result.value)
        if (result.status == Office.AsyncResultStatus.Failed) {
            write(result.error.message)
        } else {
            var attachmentContentType = item.attachments[0].contentType
            if (attachmentContentType == "application/irmaseal") {
                enableSenderinfo(item.sender.emailAddress)
                enablePolicyInfo(item.to[0].emailAddress)

                document.getElementById("run").hidden = false

                write("IRMASeal encrypted email, able to decrypt.")
                console.log("IRMASeal email")
            } else {
                console.log("No IRMASeal email")
                write("No IRMASeal email, cannot decrypt.")
                document.getElementById("run").hidden = false
            }
        }
    })
}

// Writes to a div with id='message' on the page.
function write(message) {
    document.getElementById("item-status").innerHTML += message
    document.getElementById("item-status").innerHTML += "<br/>"
}

function enablePolicyInfo(receiver: string) {
    document.getElementById("item-policy").hidden = false
    document.getElementById("item-policy").innerHTML = receiver
}

function enableSenderinfo(sender: string) {
    document.getElementById("item-sender").hidden = false
    document.getElementById("item-sender").innerHTML += sender
}

function writeMail(message) {
    document.getElementById("decrypted-text").innerHTML = message
}

function getGraphAPIToken() {
    mailbox.getCallbackTokenAsync(graphAPITokenCallback)
}

function successMessageReceived(returnData) {
    const identity = mailbox.userProfile.emailAddress
    console.log("current identity: ", identity)

    console.log("MIME: ", returnData)

    const readMail = new ReadMail()
    readMail.parseMail(returnData)

    console.log("Version: ", readMail.getVersion())

    const input = readMail.getCiphertext()

    Client.build("https://irmacrypt.nl/pkg").then((client) => {
        const readableStream = client.createUint8ArrayReadable(input)
        client
            .extractMetadata(readableStream)
            .then(async ({ metadata, header }) => {
                console.log("metadata extract", metadata)
                const {
                    identity: { attribute: irmaIdentity, timestamp: timestamp },
                } = metadata.to_json()

                const metajson = metadata.to_json()

                var session = client.createPKGSession(irmaIdentity, timestamp)

                var irma = new IrmaCore({ debugging: true, session: session })
                irma.use(IrmaClient)
                irma.use(IrmaPopup)

                const usk = await irma.start()
                const keys = metadata.derive_keys(usk)
                const iv = metajson.iv
                const decrypt = true

                console.log("Input length: ", input.byteLength)

                const plainBytes: Uint8Array = await client.symcrypt({
                    keys,
                    iv,
                    header,
                    input,
                    decrypt,
                })

                const mail: string = new TextDecoder().decode(plainBytes)
                console.log("Mail content: ", mail)

                // decrypt attachments
                const attachments = readMail.getAttachments()
                for (let i = 0; i < attachments.length; i++) {
                    const attachment = attachments[i]
                    const iv = attachment.nonce
                    const input = attachment.body

                    // decrypt only attachments that end on .enc (as we need to skip the version and CT part)
                    const attachmentBytes: Uint8Array = await client.symcrypt({
                        keys,
                        iv,
                        header,
                        input,
                        decrypt,
                    })

                    const base64EncodedAttachment: string = new TextDecoder().decode(
                        attachmentBytes
                    )

                    document.getElementById("attachments").style.display =
                        "flex"

                    // create for each attachment a "div" element, which we assign a click event, and the data as a blob object via jQueries data storage.
                    // why to use blob (uint8array) instead of base64 encoded string: https://blobfolio.com/2019/better-binary-batter-mixing-base64-and-uint8array/
                    // when the user clicks, the blob is attached to a temporary anchor element and triggered programmatically to download the file.
                    const a = document
                        .getElementById("attachmentList")
                        .appendChild(document.createElement("div"))
                    a.innerHTML = attachment.fileName
                    a.onclick = downloadBlobHandler
                    $(a).data("blob", base64toBlob(base64EncodedAttachment))
                }

                document.getElementById("decryptinfo").style.display = "none"

                document.getElementById("irmaapp").style.display = "none"

                document.getElementById("bg_decrypted_txt").style.display =
                    "block"

                document.getElementById("idlock_svg_decrypt").style.display =
                    "block"

                document.getElementById("idlock_svg").style.display = "none"

                document.getElementById("expires").style.display = "none"

                document.getElementById("irma-popup-web-form").style.display =
                    "none"

                document.getElementById("info_message_text").innerHTML =
                    "Decrypted message from"

                writeMail(mail)
            })
    })
}

function graphAPITokenCallback(asyncResult) {
    if (asyncResult.status === "succeeded") {
        var restHost = Office.context.mailbox.restUrl
        var getMessageUrl =
            restHost + "/v2.0/me/messages/" + getItemRestId() + "/$value"

        console.log("Try to receive MIME")

        $.ajax({
            url: getMessageUrl,
            headers: { Authorization: "Bearer " + asyncResult.value },
            success: successMessageReceived,
            // eslint-disable-next-line no-unused-vars
            error: function (xhr, _status, _error) {
                var errorMessage = xhr.status + ": " + xhr.statusText
                console.log("Error - " + errorMessage)
            },
        })
    } else {
        console.log(
            "Could not get callback token: " + asyncResult.error.message
        )
    }
}

function getItemRestId() {
    if (Office.context.mailbox.diagnostics.hostName === "OutlookIOS") {
        // itemId is already REST-formatted.
        return Office.context.mailbox.item.itemId
    } else {
        // Convert to an item ID for API v2.0.
        return Office.context.mailbox.convertToRestId(
            Office.context.mailbox.item.itemId,
            Office.MailboxEnums.RestVersion.v2_0
        )
    }
}

// helper functions for attachment conversion and download

const base64toBlob = function (data: string) {
    const contentType = "application/octet-stream"
    const buff = Buffer.from(data, "base64")
    return new Blob([buff.buffer], { type: contentType })
}

const downloadBlobAsFile = function (blob: Blob, filename: string) {
    const contentType = "application/octet-stream"
    if (!blob) {
        console.error("No data")
        return
    }

    const a = document.createElement("a")
    a.download = filename
    a.href = window.URL.createObjectURL(blob)
    a.dataset.downloadurl = [contentType, a.download, a.href].join(":")

    const e = new MouseEvent("click")
    a.dispatchEvent(e)
}

function downloadBlobHandler(e) {
    const target = e.target
    const filename = target.innerHTML
    const data = $(target).data("blob")
    downloadBlobAsFile(data, filename)
}
