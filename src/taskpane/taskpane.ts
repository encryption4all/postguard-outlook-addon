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

var BOUNDARY: string

function extractFromMime(dataBuffer: string): string {
    // First extract boundary
    BOUNDARY = dataBuffer
        .match(/boundary=(.*)/gm)[0]
        .replace("boundary=", "")
        .replace(/"/g, "")

    console.log("Boundary: ", BOUNDARY)

    const [_, section2, section3, section4] = dataBuffer
        .split(`--${BOUNDARY}`)
        .slice(0, -1)

    var ctPart: string
    var versionPart: string
    if (
        section4 === undefined ||
        section3.search("Content-Type: application/octet-stream") !== -1
    ) {
        versionPart = section2
        ctPart = section3
    } else {
        versionPart = section3
        ctPart = section4
    }

    const sec2RegExp = /Content-Type: application\/irmaseal\r?\n(.*)\r?\n/
    const sec3RegExp = /Content-Transfer-Encoding: base64\r?\n\r?\n([\s\S]*)/gm

    versionPart = versionPart
        .match(sec3RegExp)[0]
        .replace(sec3RegExp, "$1")
        .replace(" ", "")
        .replace("\r\n", "")
    const version = Buffer.from(versionPart, "base64").toString("utf-8")

    const bytes = ctPart
        .match(sec3RegExp)[0]
        .replace(sec3RegExp, "$1")
        .replace(/(?:\r\n|\r|\n| )/g, "")

    console.log(version)
    console.log("Bytes: ", bytes)

    return bytes
}

function successMessageReceived(returnData) {
    var identity = mailbox.userProfile.emailAddress
    console.log("current identity: ", identity)

    console.log("MIME: ", returnData)

    const bytes = extractFromMime(returnData)

    const sealBytes: Uint8Array = new Uint8Array(Buffer.from(bytes, "base64"))
    console.log("ct bytes: ", sealBytes)

    const readableStream = createUint8ArrayReadable(sealBytes)

    Client.build("https://irmacrypt.nl/pkg").then((client) => {
        console.log("Client build")
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
                        ).style.display = "block"

                        document.getElementById(
                            "idlock_svg_decrypt"
                        ).style.display = "block"

                        document.getElementById("idlock_svg").style.display =
                            "none"

                        document.getElementById("expires").style.display =
                            "none"

                        document.getElementById("info_message_text").innerHTML =
                            "Decrypted message from"

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
