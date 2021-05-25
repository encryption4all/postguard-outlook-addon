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
    Attribute,
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
        document.getElementById("app-body").style.display = "flex"
        document.getElementById("run").onclick = run
    }
})

var item
var mailbox

export async function run() {
    // Write message property value to the task pane
    console.log("Run method")
    $(function () {
        setItemBody()
    })
}

Office.initialize = function () {
    item = Office.context.mailbox.item
    mailbox = Office.context.mailbox

    console.log("Initialize")

    /*import("@e4a/irmaseal-wasm-bindings/index.js").then((wasm) => {
        let myBigThing = new wasm.MyBigThing()
        console.log("wasm ", wasm)
        console.log(myBigThing.greet())
        
    })*/
}

// Get the body type of the composed item, and set data in
// in the appropriate data type in the item body.
function setItemBody() {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
        console.log(result.value)
        if (result.status == Office.AsyncResultStatus.Failed) {
            write(result.error.message)
        } else {
            var attachmentContentType = item.attachments[0].contentType
            // Successfully got the type of item body.
            // Set data of the appropriate type in body.
            if (attachmentContentType == "application/irmaseal") {
                console.log("IRMASeal email")
                // Body is of text type.
                write(
                    "This is an IRMASeal encrypted email, starting decrypting process ..."
                )
                getGraphAPIToken() //AttachmentToken()
            } else {
                console.log("No IRMASeal email")
            }
        }
    })
}

// Writes to a div with id='message' on the page.
function write(message) {
    document.getElementById("item-subject").innerHTML += message
}

function writeMail(message) {
    document.getElementById("item-mail").innerHTML += message
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

const client2: Client = await Client.build(
    "https://irmacrypt.nl/pkg",
    true,
    Office.context.roamingSettings
)

function successMessageReceived(returnData) {
    var identity = mailbox.userProfile.emailAddress
    console.log("current identity: ", identity)

    const bytes = extractFromMime(returnData)
    const sealBytes: Uint8Array = new Uint8Array(Buffer.from(bytes, "base64"))
    console.log("ct bytes: ", bytes)

    const readableStream = createUint8ArrayReadable(sealBytes)

    Client.build(
        "https://irmacrypt.nl/pkg",
        true,
        Office.context.roamingSettings
    ).then((client) => {
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

    /*const attribute: Attribute = {
        type: "pbdf.sidn-pbdf.email.email",
        value: identity,
    }

    const meta = client.createMetadata(attribute)
    console.log("Meta: ", meta)
    */

    //.finally(() => window.close())
    //});

    //.catch((err) => console.log(err));
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
