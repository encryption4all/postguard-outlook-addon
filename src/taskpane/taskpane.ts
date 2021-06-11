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

import {
    Client,
    createUint8ArrayReadable,
    KeySet,
    symcrypt,
    MetadataReaderResult,
} from "@e4a/irmaseal-client"

var Buffer = require("buffer/").Buffer

/* global document, Office */

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
    document.getElementById("decrypted-text").innerHTML += message
}

function getGraphAPIToken() {
    mailbox.getCallbackTokenAsync(graphAPITokenCallback)
}

const BOUNDARY = "foo"

function extractFromMime(dataBuffer: string): string {
    const [section1, section2, section3] = dataBuffer
        .split(`--${BOUNDARY}`)
        .slice(0, -1)

    const sec1RegExp = /(.*)\r?\n--foo/
    const sec2RegExp = /Content-Type: application\/irmaseal\r?\nVersion: (.*)\r?\n/
    const sec3RegExp = /Content-Type: application\/octet-stream\r?\n(.*)\r?\n/

    const plain = section1.replace(sec1RegExp, "$1")
    const version = section2.replace(sec2RegExp, "$1")
    const bytes = section3
        .replace(sec3RegExp, "$1")
        .replace(" ", "")
        .replace("\n", "")

    // console.log(`info: ${version},\n bytes: ${bytes}`)

    // return { bytes, version } //output
    return bytes //output
}

function successMessageReceived(returnData) {
    var identity = mailbox.userProfile.emailAddress
    console.log("current identity: ", identity)

    const bytes = extractFromMime(returnData)
    const sealBytes: Uint8Array = new Uint8Array(Buffer.from(bytes, "base64"))
    console.log("ct bytes: ", bytes)

    const readableStream = createUint8ArrayReadable(sealBytes)

    Client.build("https://irmacrypt.nl/pkg").then((client) => {
        client
            .extractMetadata(readableStream)
            .then((metadata: MetadataReaderResult) => {
                console.log("metadata extract", metadata)
                let metajson = metadata.metadata.to_json()
                console.log("metadata to json: ", metajson)

                client
                    .requestToken(metajson.identity.attribute)
                    .then((token) =>
                        client.requestKey(token, metajson.identity.timestamp)
                    )
                    .then(async (usk) => {
                        const keys: KeySet = metadata.metadata.derive_keys(usk)
                        const plainBytes: Uint8Array = await symcrypt(
                            keys,
                            metajson.iv,
                            metadata.header,
                            sealBytes,
                            true
                        )
                        const mail: string = new TextDecoder().decode(
                            plainBytes
                        )
                        console.log("Mail content: ", mail)
                        document.getElementById("decryptinfo").style.display =
                            "none"
                        document.getElementById("irmaapp").style.display =
                            "none"
                        document.getElementById(
                            "bg_decrypted_txt"
                        ).hidden = false
                        writeMail(mail)
                    })
                    .catch((err) => {
                        console.log("Error decrypting mail: ", err, err.stack)
                    })
            })
            .catch((err) => {
                console.error("Error exracting metadata: ", err, err.stack)
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
            error: function (xhr, status, error) {
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
