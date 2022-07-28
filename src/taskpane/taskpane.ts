/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global $, Office */

// images references in the manifest
import '../../assets/16.png'
import '../../assets/32.png'
import '../../assets/80.png'

import 'web-streams-polyfill'

import { getGlobal, getItemRestId, isPostGuardEmail } from '../helpers/utils'
import { successMailReceived } from '../decryptdialog/decrypt'

const getLogger = require('webpack-log')
const decryptLog = getLogger({ name: 'PostGuard decrypt log' })

var item: Office.MessageRead

const g = getGlobal() as any

/**
 * onReady function called when file is initialized
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById('sideload-msg').style.display = 'none'
    document.getElementById('app-body').hidden = false
    item = Office.context.mailbox.item
    $(function () {
      if (isPostGuardEmail()) {
        getGraphAPIToken()
        enableSenderinfo(item.sender.emailAddress)
      } else {
        write('No Postguard email, cannot decrypt.')
      }
    })
  }
})

/**
 * Shows a message in the taskpane, and disables all other elements
 * @param message Message to be displayed
 */
function write(message) {
  decryptLog.warn(message)
  document.getElementById('info_message').style.display = 'none'
  document.getElementById('decryptinfo').style.display = 'none'
  document.getElementById('irmaapp').style.display = 'none'
  document.getElementById('header_text').style.display = 'none'
  document.getElementById('decrypted').style.display = 'none'
  document.getElementById('loading').style.display = 'none'
  document.getElementById('status-container').hidden = false
  document.getElementById('status').innerHTML = message
}

/**
 * Enables sender information
 * @param sender The sender of the mail
 */
function enableSenderinfo(sender: string) {
  document.getElementById('item-sender').hidden = false
  document.getElementById('item-sender').innerHTML += sender
}

/**
 * Callback from graph API token request
 * @param token MS Graph API authentication token
 */
async function graphAPITokenCallback(token) {
  var getMessageUrl =
    'https://graph.microsoft.com/v1.0/me/messages/' +
    getItemRestId() +
    '/$value'

  decryptLog.info('Try to receive MIME')

  try {
    const mime = await $.ajax({
      url: getMessageUrl,
      headers: { Authorization: 'Bearer ' + token }
    })

    g.token = token
    g.recipient = Office.context.mailbox.userProfile.emailAddress
    g.mailId = item.itemId
    g.attachmentId = item.attachments[0].id
    g.msgFunc = write

    await successMailReceived(mime)
  } catch (error) {
    console.error(error)
  }
}

/**
 * Initializes dialog for authentication to Graph API
 */
async function getGraphAPIToken() {
  showLoginPopup('/fallbackauthdialog.html')
}

var loginDialog

/**
 * This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
 * @param arg The message received
 */
async function processMessage(arg) {
  let messageFromDialog = JSON.parse(arg.message)

  if (messageFromDialog.status === 'success') {
    // We now have a valid access token.
    loginDialog.close()
    console.log('Valid token: ', JSON.stringify(messageFromDialog.result))
    console.log('Logginger: ', JSON.stringify(messageFromDialog.logging))
    graphAPITokenCallback(messageFromDialog.result.accessToken)
  } else {
    // Something went wrong with authentication or the authorization of the web application.
    loginDialog.close()
    console.log(
      'Message from dialog error: ',
      JSON.stringify(messageFromDialog.error.toString())
    )
  }
}

/**
 * Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
 * @param url The URL pointing to the dialog / pop-up
 */
function showLoginPopup(url) {
  var fullUrl =
    location.protocol +
    '//' +
    location.hostname +
    (location.port ? ':' + location.port : '') +
    url +
    '?currentAccountMail=' +
    Office.context.mailbox.userProfile.emailAddress

  // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
  Office.context.ui.displayDialogAsync(
    fullUrl,
    { height: 60, width: 30 },
    function (result) {
      console.log('Dialog has initialized. Wiring up events')
      loginDialog = result.value
      loginDialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        processMessage
      )
    }
  )
}
