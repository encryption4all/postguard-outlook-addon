/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import '../../assets/icon-16.png'
import '../../assets/icon-32.png'
import '../../assets/icon-80.png'

import 'web-streams-polyfill'

import { ReadMail } from '@e4a/irmaseal-mail-utils/dist/index'

import * as IrmaCore from '@privacybydesign/irma-core'
import * as IrmaClient from '@privacybydesign/irma-client'
import * as IrmaWeb from '@privacybydesign/irma-web'
import {
  getItemRestId,
  hashString,
  IAttachmentContent,
  isPostGuardEmail,
  replaceMailBody
} from '../helpers/utils'
import jwtDecode, { JwtPayload } from 'jwt-decode'

const getLogger = require('webpack-log')
const log = getLogger({ name: 'PostGuard decrypt log' })

const mod_promise = import('@e4a/irmaseal-wasm-bindings')
const mod = await mod_promise
const simpleParser = require('mailparser').simpleParser

const hostname = 'https://main.irmaseal-pkg.ihub.ru.nl'
const email_attribute = 'pbdf.sidn-pbdf.email.email'

/* global $, Office */
var item: Office.MessageRead
var mailbox: Office.Mailbox

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById('sideload-msg').style.display = 'none'
    document.getElementById('app-body').hidden = false
    item = Office.context.mailbox.item
    mailbox = Office.context.mailbox
    $(function () {
      if (isPostGuardEmail()) {
        getGraphAPIToken()
        enableSenderinfo(item.sender.emailAddress)
        enablePolicyInfo(item.to[0].emailAddress)
      } else {
        write('No Postguard email, cannot decrypt.')
      }
    })
  }
})

async function successMessageReceived(mime: string, token: string) {
  const recipient_id = mailbox.userProfile.emailAddress
  console.log('current identity: ', recipient_id)
  const conjunction = [{ t: email_attribute, v: recipient_id }]
  const hashConjunction = await hashString(JSON.stringify(conjunction))

  const readMail = new ReadMail()
  readMail.parseMail(mime)
  const input = readMail.getCiphertext()
  const readable: ReadableStream = newReadableStreamFromArray(input)

  const unsealer = await mod.Unsealer.new(readable)
  const hidden = unsealer.get_hidden_policies()

  document.getElementById('qrcodecontainer').style.display = 'block'

  let localJwt = window.localStorage.getItem(`jwt_${hashConjunction}`)

  // if JWT in local storage is null, we need to execute IRMA disclosure session
  if (localJwt === null) {
    log.info('JWT not stored within localStorage.')
    localJwt = await executeIrmaDisclosureSession(conjunction, hashConjunction)
  }

  const decoded = jwtDecode<JwtPayload>(localJwt)
  // if JWT is expired, create new IRMA session
  if (Date.now() / 1000 > decoded.exp) {
    log.info('JWT expired.')
    localJwt = await executeIrmaDisclosureSession(conjunction, hashConjunction)
  }

  // retrieve USK
  const keyResp = await $.ajax({
    url: `${hostname}/v2/request/key/${hidden[recipient_id].ts.toString()}`,
    headers: { Authorization: 'Bearer ' + localJwt }
  })

  if (keyResp.status !== 'DONE' || keyResp.proofStatus !== 'VALID') {
    log.error('JWT invalid or IRMA session not done.')
  } else {
    log.info('JWT valid, continue.')

    let plain = new Uint8Array(0)
    const writable = new WritableStream({
      write(chunk) {
        plain = new Uint8Array([...plain, ...chunk])
      }
    })

    await unsealer.unseal(recipient_id, keyResp.key, writable)
    const mail: string = new TextDecoder().decode(plain)

    // Parse inner mail via simpleParser
    let parsed = await simpleParser(mail)
    showMailContent(parsed.html)
    showAttachments(parsed.attachments)

    // store mail id such that on send feature can check if mail was encrypted before
    // Bernard suggest to add custom header to decrypted mail, but cannot do so in Outlook ..
    const itemId = getItemRestId()
    window.localStorage.setItem(`mailid_${itemId}`, 'true')

    const attachments: IAttachmentContent[] = parsed.attachments.map(
      (attachment) => {
        const attachmentContent = Buffer.from(attachment.content).toString(
          'base64'
        )
        return {
          filename: attachment.filename,
          content: attachmentContent,
          isInline: false
        }
      }
    )

    if (Office.context.requirements.isSetSupported('DialogApi', '1.2')) {
      log.info('Dialog API 1.2 supported')
    }

    replaceMailBody(token, mailbox.item, parsed.html, attachments)
  }
}

async function executeIrmaDisclosureSession(
  conjunction: object,
  hashConjunction: string
) {
  // calculate diff in seconds between now and tomorrow 4 am
  let tomorrow = new Date()
  tomorrow.setDate(tomorrow.getDate() + 1)
  tomorrow.setHours(4, 0, 0, 0)
  const now = new Date()
  const seconds = Math.floor((tomorrow.getTime() - now.getTime()) / 1000)
  console.log('Diff in seconds: ', seconds)

  const requestBody = {
    con: conjunction,
    validity: seconds
  }

  const language = Office.context.displayLanguage.toLowerCase().startsWith('nl')
    ? 'nl'
    : 'en'

  const irma = new IrmaCore({
    debugging: true,
    element: '#qrcode',
    language: language,
    session: {
      url: hostname,
      start: {
        url: (o) => `${o.url}/v2/request/start`,
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestBody)
      },
      mapping: {
        sessionPtr: (r) => {
          const ptr = r.sessionPtr
          ptr.u = `https://ihub.ru.nl/irma/1/${ptr.u}`
          return ptr
        }
      },
      result: {
        url: (o, { sessionToken: token }) => `${o.url}/v2/request/jwt/${token}`,
        parseResponse: (r) => {
          return new Promise((resolve, reject) => {
            if (r.status != '200') reject('not ok')
            resolve(r.url)
          })
        }
      }
    }
  })

  irma.use(IrmaClient)
  irma.use(IrmaWeb)
  // disclose and retrieve JWT URL
  const jwtUrl = await irma.start()
  // retrieve JWT, add to local storage, and return
  const localJwt = await $.ajax({ url: jwtUrl })
  window.localStorage.setItem(`jwt_${hashConjunction}`, localJwt)
  return localJwt
}

async function graphAPITokenCallback(token) {
  var getMessageUrl =
    'https://graph.microsoft.com/v1.0/me/messages/' +
    getItemRestId() +
    '/$value'

  console.log('Try to receive MIME')

  try {
    const mime = await $.ajax({
      url: getMessageUrl,
      headers: { Authorization: 'Bearer ' + token }
    })
    await successMessageReceived(mime, token)
  } catch (error) {
    console.error(error)
  }
}

function newReadableStreamFromArray(array) {
  return new ReadableStream({
    start(controller) {
      controller.enqueue(array)
      controller.close()
    }
  })
}

const downloadBlobAsFile = function (blob: Blob, filename: string) {
  const contentType = 'application/octet-stream'
  if (!blob) {
    console.error('No data')
    return
  }

  const a = document.createElement('a')
  a.download = filename
  a.href = window.URL.createObjectURL(blob)
  a.dataset.downloadurl = [contentType, a.download, a.href].join(':')

  const e = new MouseEvent('click')
  a.dispatchEvent(e)
}

function downloadBlobHandler(e) {
  const target = e.target
  const filename = target.innerHTML
  const data = $(target).data('blob')
  downloadBlobAsFile(data, filename)
}

function write(message) {
  document.getElementById('info_message').style.display = 'none'
  document.getElementById('decryptinfo').style.display = 'none'
  document.getElementById('irmaapp').style.display = 'none'
  document.getElementById('header_text').style.display = 'none'
  document.getElementById('status-container').hidden = false
  document.getElementById('status').innerHTML = message
}

function enablePolicyInfo(receiver: string) {
  document.getElementById('item-policy').hidden = false
  document.getElementById('item-policy').innerHTML = receiver
}

function enableSenderinfo(sender: string) {
  document.getElementById('item-sender').hidden = false
  document.getElementById('item-sender').innerHTML += sender
}

function showMailContent(message) {
  document.getElementById('decryptinfo').style.display = 'none'
  document.getElementById('irmaapp').style.display = 'none'
  document.getElementById('idlock_svg').style.display = 'none'
  document.getElementById('header_text').style.display = 'none'

  document.getElementById('bg_decrypted_txt').style.display = 'block'
  document.getElementById('idlock_svg_decrypt').style.display = 'block'

  document.getElementById('info_message_text').innerHTML =
    'Decrypted message from'
  document.getElementById('decrypted-text').innerHTML = message
}

function showAttachments(attachments) {
  for (let i = 0; i < attachments.length; i++) {
    document.getElementById('attachments').style.display = 'flex'
    // create for each attachment a "div" element, which we assign a click event, and the data as a blob object via jQueries data storage.
    // why to use blob (uint8array) instead of base64 encoded string: https://blobfolio.com/2019/better-binary-batter-mixing-base64-and-uint8array/
    // when the user clicks, the blob is attached to a temporary anchor element and triggered programmatically to download the file.
    const attachment = attachments[i]
    const blob = new Blob([attachment.content.buffer], {
      type: attachment.contentType
    })

    const a = document
      .getElementById('attachmentList')
      .appendChild(document.createElement('div'))
    a.innerHTML = attachment.filename
    a.onclick = downloadBlobHandler
    $(a).data('blob', blob)
  }
}

async function getGraphAPIToken() {
  showLoginPopup('/fallbackauthdialog.html')
}

var loginDialog

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
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

// Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
function showLoginPopup(url) {
  var fullUrl =
    location.protocol +
    '//' +
    location.hostname +
    (location.port ? ':' + location.port : '') +
    url

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
