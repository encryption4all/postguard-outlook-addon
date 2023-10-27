/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to use MSAL.js to get an access token to Microsoft Graph an pass it to the task pane.
 */

/* global $, Office */

// images references in the manifest
import '../../assets/16.png'
import '../../assets/32.png'
import '../../assets/80.png'

import 'web-streams-polyfill'

import { ReadMail } from '@e4a/irmaseal-mail-utils/dist/readMail'

import * as YiviCore from '@privacybydesign/yivi-core'
import * as YiviClient from '@privacybydesign/yivi-client'
import * as YiviWeb from '@privacybydesign/yivi-web'
import {
  _getMobileUrl,
  checkLocalStorage,
  getGlobal,
  getPostGuardHeaders,
  getDecryptionUSK,
  hashCon,
  htmlBodyType,
  IAttachmentContent,
  newReadableStreamFromArray,
  PKG_URL,
  removeAttachment,
  type_to_image,
  getPublicKey,
  storeLocalStorage
} from '../helpers/utils'
import sanitizeHtml from 'sanitize-html'

import i18next from 'i18next'
import translationEN from '../../locales/en.json'
import translationNL from '../../locales/nl.json'

// eslint-disable-next-line no-undef
const getLogger = require('webpack-log')
const decryptLog = getLogger({ name: 'PostGuard decrypt log' })

const mod_promise = require('@e4a/pg-wasm')

const vk_promise: Promise<string> = getPublicKey(true)

const [vk, mod] = await Promise.all([vk_promise, mod_promise])

// eslint-disable-next-line no-undef
const simpleParser = require('mailparser').simpleParser

const EMAIL_ATTRIBUTE_TYPE = 'pbdf.sidn-pbdf.email.email'

// eslint-disable-next-line no-undef
var Buffer = require('buffer/').Buffer

const g = getGlobal() as any

/**
 * Initialization function which also extracts the URL params
 */
Office.initialize = function () {
  Office.onReady(() => {
    if (Office.context.mailbox === undefined) {
      decryptLog.info('Decrypt dialog openend!')
      const urlParams = new URLSearchParams(window.location.search)
      g.token = Buffer.from(urlParams.get('token'), 'base64').toString('utf-8')
      g.recipient = urlParams.get('recipient').toLowerCase()
      g.mailId = urlParams.get('mailid')
      g.attachmentId = urlParams.get('attachmentid')
      g.msgFunc = passMsgToParent
      g.sender = urlParams.get('sender')

      i18next.init({
        lng: Office.context.displayLanguage.toLowerCase().startsWith('nl')
          ? 'nl'
          : 'en',
        debug: true,
        resources: {
          en: {
            translation: translationEN
          },
          nl: {
            translation: translationNL
          }
        }
      })

      $(function () {
        getMailObject()
      })
    }
  })
}

/**
 * Passes a message to the parent
 * TODO: Make msg object to be able to pass status (if error then close dialog in parent)
 * @param msg The message
 */
function passMsgToParent(msg: string) {
  if (Office.context.mailbox === undefined) {
    Office.context.ui.messageParent(msg)
  }
}

/**
 * Handles an jQuery ajax error
 * @param $xhr The error
 */
function handleAjaxError($xhr) {
  var data = $xhr.responseJSON
  decryptLog.error('Ajax error: ', data)
  const msg =
    'Error during decryption, please try again or contact your administrator.'
  g.msgFunc(msg)
}

/**
 * Gets the mail object as MIME
 */
function getMailObject() {
  var getMessageUrl =
    'https://graph.microsoft.com/v1.0/me/messages/' + g.mailId + '/$value'

  decryptLog.info('Try to receive MIME via ', getMessageUrl)

  fetch(getMessageUrl, {
    headers: new Headers({
      Authorization: 'Bearer ' + g.token
    })
  })
    .then((response) => {
      if (response.ok) {
        return response.text()
      }
      throw new Error('Something went wrong when tryng to get MIME')
    })
    .then(successMailReceived)
    // eslint-disable-next-line no-unused-vars
    .catch((err) => {
      decryptLog.error(err)
      passMsgToParent(
        'Error during decryption, please try again or contact your administrator (' +
          err +
          ')'
      )
    })
}

/**
 * Handling decryption of the mail after it has been received
 * @param mime The mime message
 */
export async function successMailReceived(mime) {
  decryptLog.info('Success MIME mail received: ', mime)

  const readMail = new ReadMail()
  readMail.parseMail(mime)
  const input = readMail.getCiphertext()
  const readable: ReadableStream = newReadableStreamFromArray(input)

  const unsealer = await mod.StreamUnsealer.new(readable, vk)
  const recipients = unsealer.inspect_header()
  const recipientId = g.recipient

  const me = recipients.get(recipientId)

  if (!me) {
    const e = new Error('recipient identifier not found in header')
    e.name = 'RecipientUnknownError'
    throw e
  }

  const keyRequest = Object.assign({}, me)
  let hints = me.con

  // Convert hints.
  hints = hints.map(({ t, v }) => {
    if (t === EMAIL_ATTRIBUTE_TYPE) return { t, v: recipientId }
    else return { t, v }
  })

  // Convert hidden policy to attribute request.
  keyRequest.con = keyRequest.con.map(({ t, v }) => {
    if (t === EMAIL_ATTRIBUTE_TYPE) return { t, v: recipientId }
    else if (v === '' || v.includes('*')) return { t }
    else return { t, v }
  })

  decryptLog.info('Hints: ', hints)
  decryptLog.info('Trying decryption with policy: ', keyRequest)

  const localJwt = await checkLocalStorage(hints).catch((e) =>
    executeIrmaDisclosureSession(hints, keyRequest.con)
  )
  decryptLog.info('LocalJwt: ' + localJwt)

  // retrieve USK
  const usk = await getDecryptionUSK(localJwt, keyRequest.ts)

  let plain = new Uint8Array(0)
  const writable = new WritableStream({
    write(chunk) {
      plain = new Uint8Array([...plain, ...chunk])
    }
  })

  try {
    const senderIdentity = await unsealer.unseal(g.recipient, usk, writable)
    console.log('Sender verification successful: ', senderIdentity)

    const privBadges = senderIdentity?.private?.con ?? []
    const badges = [...senderIdentity.public.con, ...privBadges].map(
      ({ t, v }) => {
        return { type: type_to_image(t), value: v }
      }
    )
    let signingString = 'Sender signed with => '
    badges.forEach(
      (element) => (signingString += element.type + ': ' + element.value)
    )

    const mail: string = new TextDecoder().decode(plain)

    // store JWT locally after unsealing successfully
    storeLocalStorage(hints, localJwt).catch((e) => decryptLog.error(e))

    // Parse inner mail via simpleParser
    let parsed = await simpleParser(mail)
    // body can be either HTML encoded, or text as HTML encoded
    const body = parsed.html ? parsed.html : parsed.textAsHtml

    // TO and CC for display purposes
    let to = ''
    if (parsed.to !== undefined) {
      to = parsed.to.value
        .map(function (to) {
          return to.address
        })
        .join(',')
    }

    let cc = ''
    if (parsed.cc !== undefined) {
      cc = parsed.cc.value
        .map(function (cc) {
          return cc.address
        })
        .join(',')
    }

    showMailContent(
      parsed.subject,
      body,
      parsed.from.value[0].address,
      to,
      cc,
      parsed.date.toLocaleString()
    )
    showAttachments(parsed.attachments)

    // prepare attachments to be added to mail via Graph API
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

    replaceMailBody(body, parsed.subject, attachments)
    passMsgToParent('Succesfully decrypted. ' + signingString)
  } catch (error) {
    if (error.name === 'OperationError') {
      g.msgFunc('Disclosed identity does not match requested policy')
    } else {
      throw error
    }
  }
}

/**
 * Replaces the mail body of the current mail
 * @param body The body of the decrypted mail
 * @param subject The subject of the decrypted mail
 * @param attachments The attachments of the decrypted mail
 */
function replaceMailBody(
  body: string,
  subject: string,
  attachments: IAttachmentContent[]
) {
  const messageUrl = `https://graph.microsoft.com/v1.0/me/messages/${g.mailId}`
  const payload = {
    body: {
      contentType: htmlBodyType,
      content: body
    },
    subject: subject
  }
  $.ajax({
    type: 'PATCH',
    contentType: 'application/json',
    url: messageUrl,
    data: JSON.stringify(payload),
    headers: {
      Authorization: 'Bearer ' + g.token
    },
    success: function (success) {
      decryptLog.info('PATCH message success: ', success)
      removeAttachment(g.token, g.mailId, g.attachmentId, attachments)
    }
  }).fail(handleAjaxError)
}

class Policy {
  t: string
  v: string
}

/**
 * Executes an IRMA disclosure session based on the policy for decryption
 * @param hints The policy with hidden values
 * @param policy The policy
 * @param sort Either Signing or Decrypting
 * @returns The JWT of the IRMA session
 */
async function executeIrmaDisclosureSession(
  hints: Policy[],
  policy: Policy[]
): Promise<string> {
  // show HTML elements needed
  document.getElementById('info_message').style.display = 'block'
  document.getElementById('header_text').style.display = 'block'
  document.getElementById('decryptinfo').style.display = 'block'
  document.getElementById('irmaapp').style.display = 'block'
  document.getElementById('qrcodecontainer').style.display = 'block'
  document.getElementById('loading').style.display = 'none'
  enableSenderinfo(g.sender)

  $.each(hints, function (_index, element) {
    const colon = element.v.length > 0 ? ':' : ''
    $('#attributes').append(
      `<tr><td class="attrtype">${i18next
        .t(element.t)
        .toLowerCase()}${colon}</td><td class="attrvalue">${element.v}</td><tr>`
    )
  })

  // calculate diff in seconds between now and tomorrow 4 am
  let tomorrow = new Date()
  tomorrow.setDate(tomorrow.getDate() + 1)
  tomorrow.setHours(4, 0, 0, 0)
  const now = new Date()
  const seconds =
    Math.floor((tomorrow.getTime() - now.getTime()) / 1000) % 86400
  decryptLog.info('Diff in seconds until 4 am: ', seconds)

  const requestBody = {
    con: policy,
    validity: seconds
  }

  const irma = new YiviCore({
    debugging: true,
    element: '#qrcode',
    language: 'en',
    state: {
      serverSentEvents: false,
      polling: {
        endpoint: 'status',
        interval: 500,
        startState: 'INITIALIZED'
      }
    },
    session: {
      url: PKG_URL,
      start: {
        url: (o) => `${o.url}/v2/request/start`,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-Postguard-Client-Version': getPostGuardHeaders()
        },
        body: JSON.stringify(requestBody)
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

  irma.use(YiviClient)
  irma.use(YiviWeb)
  // disclose and retrieve JWT URL
  const jwtUrl = await irma.start()
  const localJwt: string = await $.ajax({ url: jwtUrl })
  return localJwt
}

/**
 * Downloading the attachment
 * @param blob The binary data of the attachment
 * @param filename The name of the attachment
 */
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

/**
 * Handler for downloading the attachment
 * @param e The event
 */
function downloadBlobHandler(e) {
  const target = e.target
  const filename = target.innerHTML
  const data = $(target).data('blob')
  downloadBlobAsFile(data, filename)
}

/**
 * Show the content of the mail
 * @param subject The subject of the mail
 * @param body The body of the mail
 */
function showMailContent(
  subject: string,
  body: string,
  from: string,
  to: string,
  cc: string,
  received: string
) {
  document.getElementById('decryptinfo').style.display = 'none'
  document.getElementById('irmaapp').style.display = 'none'
  document.getElementById('header_text').style.display = 'none'
  document.getElementById('info_message').style.display = 'none'
  document.getElementById('loading').style.display = 'none'

  document.getElementById('app-body').style.backgroundColor = '#D7E4E9'
  document.getElementById('center').className = 'leftAndMargin'
  document.getElementById('decrypted').style.display = 'block'
  document.getElementById('decrypted_subject').innerHTML = subject
  document.getElementById('decrypted_from').innerHTML += from

  const sanitizeBody = sanitizeHtml(body)
  document.getElementById('decrypted_text').innerHTML = sanitizeBody

  if (to.length > 0) document.getElementById('decrypted_to').innerHTML += to
  else document.getElementById('decrypted_to').style.display = 'none'

  if (cc.length > 0) document.getElementById('decrypted_cc').innerHTML += cc
  else document.getElementById('decrypted_cc').style.display = 'none'

  document.getElementById('decrypted_received').innerHTML += received
}

/**
 * Show the attachments
 * @param attachments The attachments
 */
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

/**
 * Enables sender information
 * @param sender The sender of the mail
 */
function enableSenderinfo(sender: string) {
  if (sender !== undefined) {
    document.getElementById('item-sender').hidden = false
    document.getElementById('item-sender').innerHTML = sender
  }
}
