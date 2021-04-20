/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

var item;

export async function run() {
  // Write message property value to the task pane
  console.log("Run method");
  //document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
  $(function() {
    setItemBody();
  });
}

Office.initialize = function () {
    item = Office.context.mailbox.item;
    console.log("Initialize");
}

// Get the body type of the composed item, and set data in 
// in the appropriate data type in the item body.
function setItemBody() {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
            console.log(result.value);
            if (result.status == Office.AsyncResultStatus.Failed){
                write(result.error.message);
            }
            else {
                var attachmentContentType = item.attachments[0].contentType;
                // Successfully got the type of item body.
                // Set data of the appropriate type in body.
                if (attachmentContentType == "application/irmaseal") {
                    console.log("IRMASeal email");
                    // Body is of text type. 
                    write('This is an IRMASeal encrypted email, please use the IRMASeal addon to decrypt');
                    //const str = atob(item.attachments[0].);
                } else {
                    console.log("No IRMASeal email");
                }
            }
        });
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('item-subject').innerHTML += message; 
}