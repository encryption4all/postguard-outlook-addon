/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png"
import "../../assets/icon-32.png"
import "../../assets/icon-80.png"

import { Client, Attribute } from "@e4a/irmaseal-client"

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

//const client = getClient()

Office.initialize = function () {
    item = Office.context.mailbox.item
    mailbox = Office.context.mailbox

    //let client = getClient()

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
                getAttachmentToken()
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

function getAttachmentToken() {
    mailbox.getCallbackTokenAsync(attachmentTokenCallback)
}

const BOUNDARY = "foo"

function decryptData(dataBuffer: string) {
    const [section1, section2, section3] = dataBuffer
        .split(`--${BOUNDARY}`)
        .slice(0, -1)

    const sec1RegExp = /(.*)\r?\n--foo/
    const sec2RegExp = /Content-Type: application\/irmaseal\r?\nVersion: (.*)\r?\n/
    const sec3RegExp = /Content-Type: application\/octet-stream\r?\n(.*)\r?\n/

    const plain = section1.replace(sec1RegExp, "$1")
    const version = section2.replace(sec2RegExp, "$1")
    const bytes = section3.replace(sec3RegExp, "$1")

    // TODO: error handling in case of no match
    //if (!section2.match(sec2RegExp)) {
    //    DEBUG_LOG('not an IRMAseal message')
    //    return
    //}

    //DEBUG_LOG(`plain: ${plain},\n info: ${version},\n bytes: ${bytes}`)

    // For now, just pass the ciphertext bytes to the frontend
    const msg = bytes //atob(bytes.replace(/[\r\n]/g, ''))

    // We need to wrap the result into a multipart/mixed message
    // TODO: can add more here
    let output = ""
    output += `Content-Type: multipart/mixed; boundary="${BOUNDARY}"\r\n\r\n`
    output += `--${BOUNDARY}\r\n`
    output += `Content-Type: text/plain\r\n\r\n`
    output += `${msg}\r\n`
    output += `--${BOUNDARY}--\r\n`

    return output
}

async function getClient() {
    try {
        var client = await Client.build(
            "https://qrona.info/pkg",
            true,
            window.localStorage
        ) //192.168.2.5:8087');
        console.log(client)
        return client
    } catch (err) {
        console.log(err)
    }
}

const client: Client = await Client.build(
    "https://qrona.info/pkg",
    true,
    window.localStorage
)

function successMessageReceived(returnData) {
    var decryptedData = decryptData(returnData)
    console.log("decrypted data: ", decryptedData)

    var identity = Office.context.mailbox.userProfile.emailAddress

    console.log("current identity: ", identity)

    //getClient().then((client) => {
    const bytes = Buffer.from(decryptedData, "base64")

    console.log("ct bytes: ", bytes)

    // const id = client.extractIdentity(bytes)
    //console.log("identity in bytestream:", id)

    const attribute: Attribute = {
        type: "pbdf.sidn-pbdf.email.email",
        value: identity,
    }

    console.log("Client ", client)

    let meta = client.createMetadata(attribute)

    console.log("Meta: ", meta)

    /*
  console.log("Created metadata: ", meta);

  let metadata = meta.to_json();
  console.log("metadata to json: ", metadata);

  client
    .requestToken(attribute)
    .then((token) => client.requestKey(token, metadata.identity.timestamp))
    .then(async (usk) => {
      //const mail = client.decrypt(usk, bytes)
      console.log(usk);
      //await browser.messageDisplayScripts.register({
      //    js: [{ code: `document.body.textContent = "${mail.body}";` }, { file: 'display.js' }],
      //})
    })
    .catch((err) => {
      console.log("error: ", err);
    });
  //.finally(() => window.close())
  //});
*/

    //.catch((err) => console.log(err));
}

function attachmentTokenCallback(asyncResult) {
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
