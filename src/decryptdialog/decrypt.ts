/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to use MSAL.js to get an access token to Microsoft Graph an pass it to the task pane.
 */

/* global $, Office */

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
  hashString,
  htmlBodyType,
  IAttachmentContent,
  newReadableStreamFromArray,
  removeAttachment
} from '../helpers/utils'
import jwtDecode, { JwtPayload } from 'jwt-decode'

// eslint-disable-next-line no-undef
const getLogger = require('webpack-log')
const decryptLog = getLogger({ name: 'PostGuard decrypt log' })

const mod_promise = import('@e4a/irmaseal-wasm-bindings')
const mod = await mod_promise
// eslint-disable-next-line no-undef
const simpleParser = require('mailparser').simpleParser

const hostname = 'https://main.irmaseal-pkg.ihub.ru.nl'
const email_attribute = 'pbdf.sidn-pbdf.email.email'

// eslint-disable-next-line no-undef
var Buffer = require('buffer/').Buffer

var token: string
var recipient: string
var mailId: string
var attachmentId: string

/**
 * Initialization function which also extracts the URL params
 */
Office.initialize = function () {
  decryptLog.info('Decrypt dialog openend!')
  const urlParams = new URLSearchParams(window.location.search)
  token = Buffer.from(urlParams.get('token'), 'base64').toString('utf-8')
  recipient = urlParams.get('recipient')
  mailId = urlParams.get('mailid')
  attachmentId = urlParams.get('attachmentid')

  $(function () {
    getMailObject()
  })
}

/**
 * Passes a message to the parent
 * @param msg The message
 */
function passMsgToParent(msg: string) {
  Office.context.ui.messageParent(msg)
}

/**
 * Handles an jQuery ajax error
 * @param $xhr The error
 */
function handleAjaxError($xhr) {
  var data = $xhr.responseJSON
  decryptLog.error('Ajax error: ', data)
  passMsgToParent(
    'Error during decryption, please try again or contact your administrator.'
  )
}

/**
 * Gets the mail object as MIME
 */
function getMailObject() {
  var getMessageUrl =
    'https://graph.microsoft.com/v1.0/me/messages/' + mailId + '/$value'

  decryptLog.info('Try to receive MIME via ', getMessageUrl)

  fetch(getMessageUrl, {
    headers: new Headers({
      Authorization: 'Bearer ' + token
    })
  })
    .then((response) => {
      if (response.ok) {
        return response.text()
      }
      throw new Error('Something went wrong')
    })
    .then(successMailReceived)
    // eslint-disable-next-line no-unused-vars
    .catch((err) => {
      decryptLog.error(err)
      passMsgToParent(
        'Error during decryption, please try again or contact your administrator.'
      )
    })
}

/**
 * Handling decryption of the mail after it has been received
 * @param mime The mime message
 */
async function successMailReceived(mime) {
  decryptLog.info('Success MIME mail received')
  const conjunction = [{ t: email_attribute, v: recipient }]
  const hashConjunction = await hashString(JSON.stringify(conjunction))

  const readMail = new ReadMail()
  readMail.parseMail(mime)
  const input = readMail.getCiphertext()
  const readable: ReadableStream = newReadableStreamFromArray(input)

  const unsealer = await mod.Unsealer.new(readable)
  const hidden = unsealer.get_hidden_policies()

  let localJwt = window.localStorage.getItem(`jwt_${hashConjunction}`)

  // if JWT in local storage is null, we need to execute IRMA disclosure session
  if (localJwt === null) {
    decryptLog.info(
      'JWT not stored within localStorage, starting IRMA session ...'
    )
    localJwt = await executeIrmaDisclosureSession(conjunction, hashConjunction)
  }

  const decoded = jwtDecode<JwtPayload>(localJwt)
  // if JWT is expired, create new IRMA session
  if (Date.now() / 1000 > decoded.exp) {
    decryptLog.info('JWT expired.')
    localJwt = await executeIrmaDisclosureSession(conjunction, hashConjunction)
  }

  // retrieve USK
  const keyResp = await $.ajax({
    url: `${hostname}/v2/request/key/${hidden[recipient].ts.toString()}`,
    headers: { Authorization: 'Bearer ' + localJwt }
  })

  if (keyResp.status !== 'DONE' || keyResp.proofStatus !== 'VALID') {
    decryptLog.error('JWT invalid or IRMA session not done.')
    passMsgToParent('IRMA session not done, please try again')
  } else {
    decryptLog.info('JWT valid, continue.')

    let plain = new Uint8Array(0)
    const writable = new WritableStream({
      write(chunk) {
        plain = new Uint8Array([...plain, ...chunk])
      }
    })

    await unsealer.unseal(recipient, keyResp.key, writable)
    const mail: string = new TextDecoder().decode(plain)

    // Parse inner mail via simpleParser
    let parsed = await simpleParser(mail)
    showMailContent(parsed.subject, parsed.html)
    showAttachments(parsed.attachments)

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

    replaceMailBody(parsed.html, parsed.subject, attachments)
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
  const messageUrl = `https://graph.microsoft.com/v1.0/me/messages/${mailId}`
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
      Authorization: 'Bearer ' + token
    },
    success: function (success) {
      decryptLog.info('PATCH message success: ', success)
      passMsgToParent('Successfully decrypted the email with PostGuard')
      removeAttachment(token, mailId, attachmentId, attachments)
    }
  }).fail(handleAjaxError)
}

/**
 * Executes an IRMA disclosure session based on the policy
 * @param policy The policy
 * @param hashPolicy The hash of the policy
 * @returns The JWT of the IRMA session
 */
async function executeIrmaDisclosureSession(
  policy: object,
  hashPolicy: string
) {
  // show HTML elements needed
  document.getElementById('info_message').style.display = 'block'
  document.getElementById('header_text').style.display = 'block'
  document.getElementById('decryptinfo').style.display = 'block'
  document.getElementById('irmaapp').style.display = 'block'
  document.getElementById('qrcodecontainer').style.display = 'block'
  document.getElementById('loading').style.display = 'none'

  // calculate diff in seconds between now and tomorrow 4 am
  let tomorrow = new Date()
  tomorrow.setDate(tomorrow.getDate() + 1)
  tomorrow.setHours(4, 0, 0, 0)
  const now = new Date()
  const seconds = Math.floor((tomorrow.getTime() - now.getTime()) / 1000)
  decryptLog.info('Diff in seconds until 4 am tomorrow: ', seconds)

  const requestBody = {
    con: policy,
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
  window.localStorage.setItem(`jwt_${hashPolicy}`, localJwt)
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
function showMailContent(subject: string, body: string) {
  document.getElementById('decryptinfo').style.display = 'none'
  document.getElementById('irmaapp').style.display = 'none'
  document.getElementById('idlock_svg').style.display = 'none'
  document.getElementById('header_text').style.display = 'none'
  document.getElementById('info_message_text').style.display = 'none'
  document.getElementById('loading').style.display = 'none'

  document.getElementById('bg_decrypted_txt').style.display = 'block'
  document.getElementById('bg_decrypted_subject').style.display = 'block'
  document.getElementById('idlock_svg_decrypt').style.display = 'block'

  document.getElementById('decrypted-subject').innerHTML = subject
  document.getElementById('decrypted-text').innerHTML = body
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
