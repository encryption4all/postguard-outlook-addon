/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, Office */

import { ComposeMail } from '@e4a/irmaseal-mail-utils/dist/index'
import { createMimeMessage } from 'mimetext'
import {
  storeMailAsPlainLocally,
  IAttachmentContent,
  setEventError,
  htmlBodyType
} from '../helpers/utils'

// eslint-disable-next-line no-undef
var Buffer = require('buffer/').Buffer

var mailboxItem
var globalEvent

const hostname = 'https://main.irmaseal-pkg.ihub.ru.nl'
const mod_promise = import('@e4a/irmaseal-wasm-bindings')

// in bytes (1024 x 1024 = 1 MB)
// const MAX_ATTACHMENT_SIZE = 1024 * 1024

Office.initialize = () => {
  Office.onReady(() => {
    mailboxItem = Office.context.mailbox.item

    delete window.alert // assures alert works
    delete window.confirm // assures confirm works
    delete window.prompt // assures prompt works
  })
}

/**
 * Entry point function.
 * @param event
 */
// eslint-disable-next-line no-unused-vars
function encrypt(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: 'Encrypting email with IRMASeal',
    icon: 'Icon.80x80',
    persistent: true
  }

  globalEvent = event

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    'action',
    message
  )

  showLoginPopup('/fallbackauthdialog.html')
}

// Get recipient mail
function getRecipientEmails(): Promise<string[]> {
  return new Promise(function (resolve, reject) {
    mailboxItem.to.getAsync((recipients) => {
      let recipientMails = new Array()
      recipients.value.forEach((recipient) => {
        recipientMails.push(recipient.emailAddress)
      })
      if (recipientMails.length !== 0) resolve(recipientMails)
      else reject('No recipient email')
    })
  })
}

// Get cc recipient mail
function getCcRecipientEmails(): Promise<string[]> {
  return new Promise(function (resolve) {
    mailboxItem.cc.getAsync((recipients) => {
      let recipientMails = new Array()
      recipients.value.forEach((recipient) => {
        recipientMails.push(recipient.emailAddress)
      })
      resolve(recipientMails)
    })
  })
}

// Get bcc recipient mail
function getBccRecipientEmails(): Promise<string[]> {
  return new Promise(function (resolve) {
    mailboxItem.bcc.getAsync((recipients) => {
      let recipientMails = new Array()
      recipients.value.forEach((recipient) => {
        recipientMails.push(recipient.emailAddress)
      })
      resolve(recipientMails)
    })
  })
}

// Gets the mail body
async function getMailBody(): Promise<string> {
  return new Promise(function (resolve, reject) {
    mailboxItem.body.getAsync(Office.CoercionType.Html, (asyncResult) => {
      const body: string = asyncResult.value
      const pattern = /<body[^>]*>((.|[\n\r])*)<\/body>/im
      const arrayMatches = pattern.exec(body)
      const mailBody = arrayMatches[1]
      if (body !== '') resolve(mailBody)
      else reject('No body in email')
    })
  })
}

// Gets the mail subject
async function getMailSubject(): Promise<string> {
  return new Promise(function (resolve, reject) {
    mailboxItem.subject.getAsync((asyncResult) => {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        reject('Subject async failed')
      } else {
        const subject: string = asyncResult.value
        if (subject !== '') resolve(subject)
        else reject('No subject in email')
      }
    })
  })
}

async function getMailAttachments(): Promise<IAttachmentContent[]> {
  return new Promise(function (resolve, reject) {
    mailboxItem.getAttachmentsAsync(async (asyncResult) => {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        reject('Attachments async failed')
      } else {
        if (asyncResult.value.length > 0) {
          let attachmentsArray = []
          let content = ''
          for (var i = 0; i < asyncResult.value.length; i++) {
            var attachment = asyncResult.value[i]
            content = await getMailAttachmentContent(attachment.id)
            attachmentsArray.push({
              filename: attachment.name,
              content: content,
              isInline: attachment.isInline
            })
          }
          resolve(attachmentsArray)
        } else {
          reject('No attachments in email')
        }
      }
    })
  })
}

async function getMailAttachmentContent(attachmentId: string): Promise<string> {
  return new Promise(function (resolve, reject) {
    mailboxItem.getAttachmentContentAsync(attachmentId, (asyncResult) => {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        reject('Attachment content async failed')
      } else {
        if (asyncResult.value.content.length > 0) {
          resolve(asyncResult.value.content)
        } else
          reject('No attachment content in attachment with id' + attachmentId)
      }
    })
  })
}

// Encrypts and sends the mail
async function encryptAndsendMail(token) {
  const response = await fetch(`${hostname}/v2/parameters`)
  const pk = await response.json()
  const [mod] = await Promise.all([mod_promise])

  const sender = Office.context.mailbox.userProfile.emailAddress
  const email_attribute = 'pbdf.sidn-pbdf.email.email'
  const timestamp = Math.round(Date.now() / 1000)

  const recipientEmails: string[] = await getRecipientEmails()

  const policies = recipientEmails.reduce((total, recipient) => {
    total[recipient] = {
      ts: timestamp,
      con: [{ t: email_attribute, v: recipient }]
    }
    return total
  }, {})

  const ccRecipientEmails: string[] = await getCcRecipientEmails()
  const ccPolicies = ccRecipientEmails.reduce((total, recipient) => {
    total[recipient] = {
      ts: timestamp,
      con: [{ t: email_attribute, v: recipient }]
    }
    return total
  }, {})

  const bccRecipientEmails: string[] = await getBccRecipientEmails()
  const bccPolicies = bccRecipientEmails.reduce((total, recipient) => {
    total[recipient] = {
      ts: timestamp,
      con: [{ t: email_attribute, v: recipient }]
    }
    return total
  }, {})

  // Also encrypt for the sender, such that the sender can later decrypt as well.
  policies[sender] = { ts: timestamp, con: [{ t: email_attribute, v: sender }] }

  const allPolicies = { ...policies, ...ccPolicies, ...bccPolicies }

  console.log('Encrypting using the following policies: ', allPolicies)

  let mailBody = await getMailBody()

  const mailSubject = await getMailSubject()
  console.log('Mail subject: ', mailSubject)

  let attachments: IAttachmentContent[]
  await getMailAttachments()
    .then((attas) => (attachments = attas))
    .catch((error) => console.log(error))

  /* 
    const client = await Client.build("https://irmacrypt.nl/pkg")
    const controller = new AbortController()
    const cryptifyApiWrapper = new CryptifyApiWrapper(
        client,
        recipientEmail,
        sender,
        "https://dellxps"
    )*/

  // Use createMimeMessage to create inner MIME mail
  const msg = createMimeMessage()
  msg.setSender(sender)
  msg.setSubject(mailSubject)

  msg.setRecipient(recipientEmails)
  ccRecipientEmails.length > 0 && msg.setCc(ccRecipientEmails)
  bccRecipientEmails.length > 0 && msg.setBcc(bccRecipientEmails)

  // ComposeMail only used for outer mail
  const composeMail = new ComposeMail()
  composeMail.setSubject(mailSubject)
  composeMail.setSender(sender)

  recipientEmails.forEach((recipientEmail) => {
    composeMail.addRecipient(recipientEmail)
  })
  ccRecipientEmails.forEach((recipientEmail) => {
    composeMail.addCcRecipient(recipientEmail)
  })
  bccRecipientEmails.forEach((recipientEmail) => {
    composeMail.addBccRecipient(recipientEmail)
  })

  let hasAttachments: boolean = false

  if (attachments !== undefined) {
    let useCryptify = false
    for (let i = 0; i < attachments.length; i++) {
      const attachment = attachments[i]
      /*const fileBlob = new Blob([attachment.content], {
                type: "application/octet-stream",
            })
            const file = new File([fileBlob], attachment.filename, {
                type: "application/octet-stream",
            })

            // if attachment is too large, ask user if it should be encrypted via Cryptify
            if (fileBlob.size > MAX_ATTACHMENT_SIZE) {
                // TODO: Add confirmation dialog (https://theofficecontext.com/2017/06/14/dialogs-in-officejs/)
                console.log(
                    `Attachment ${attachment.filename} larger than 1 MB`
                )
                useCryptify = true
                const downloadUrl = await cryptifyApiWrapper.encryptAndUploadFile(
                    file,
                    controller
                )
                mailBody += `<p><a href="${downloadUrl}">Download ${attachment.filename} via Cryptify</a></p>`
            }
            */

      if (!attachment.isInline) {
        hasAttachments = true
        if (!useCryptify) {
          const input = new TextEncoder().encode(attachment.content)
          console.log('Attachment bytes length: ', input.byteLength)
          msg.setAttachment(
            attachment.filename,
            'application/octet-stream',
            attachment.content
          )
        }
      } else {
        // replace inline image in body
        const imageContentIDToReplace = `cid:${attachment.filename}@.*"`
        const regex = new RegExp(imageContentIDToReplace, 'g')
        mailBody = mailBody.replace(
          regex,
          `data:image;base64,${attachment.content}"`
          //attachment.filename + '"'
        )
      }
    }
  }

  let jsonInnerMail = {
    sender: { emailAddress: { address: sender } },
    subject: mailSubject,
    body: {
      contentType: htmlBodyType,
      content: mailBody
    },
    toRecipients: recipientEmails.map((recipient) => {
      return {
        emailAddress: { address: recipient }
      }
    }),
    ccRecipients: ccRecipientEmails.map((recipient) => {
      return {
        emailAddress: { address: recipient }
      }
    }),
    bccRecipients: bccRecipientEmails.map((recipient) => {
      return {
        emailAddress: { address: recipient }
      }
    }),
    hasAttachments: hasAttachments
  }

  console.log('jsonInnerMail: ', JSON.stringify(jsonInnerMail))

  console.log('Mailbody: ', mailBody)
  msg.setMessage('text/html', mailBody)

  // encrypt inner MIME mail
  const innerMail = msg.asRaw()
  const plainBytes: Uint8Array = new TextEncoder().encode(innerMail)
  const readable = new_readable_stream_from_array(plainBytes)
  let ct = new Uint8Array(0)
  const writable = new WritableStream({
    write(chunk) {
      ct = new Uint8Array([...ct, ...chunk])
    }
  })

  await mod.seal(pk.publicKey, policies, readable, writable)
  console.log('ct: ', ct)

  composeMail.setPayload(ct)

  // get outer mail to send email via Graph API
  const outerMail = composeMail.getMimeMail()
  const message = Buffer.from(outerMail).toString('base64')
  const sendMessageUrl = 'https://graph.microsoft.com/v1.0/me/sendMail'
  console.log('Trying to send email via ', sendMessageUrl)

  $.ajax({
    type: 'POST',
    contentType: 'text/plain',
    url: sendMessageUrl,
    data: message,
    headers: {
      Authorization: 'Bearer ' + token
    },
    success: function (success) {
      console.log('Sendmail success: ', success)

      const successMsg: Office.NotificationMessageDetails = {
        type: Office.MailboxEnums.ItemNotificationMessageType
          .InformationalMessage,
        message: 'Successfully encrypted and send email',
        icon: 'Icon.80x80',
        persistent: true
      }

      storeMailAsPlainLocally(token, jsonInnerMail, attachments, 'CryptifySend')

      Office.context.mailbox.item.notificationMessages.replaceAsync(
        'action',
        successMsg
      )

      globalEvent.completed()
    }
  }).fail(function ($xhr) {
    var data = $xhr.responseJSON
    console.log('Ajax error: ', data)
    setEventError()
  })
}

function new_readable_stream_from_array(array) {
  return new ReadableStream({
    start(controller) {
      controller.enqueue(array)
      controller.close()
    }
  })
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
    encryptAndsendMail(messageFromDialog.result.accessToken)
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

function getGlobal() {
  return typeof self !== 'undefined'
    ? self
    : typeof window !== 'undefined'
    ? window
    : typeof global !== 'undefined'
    ? // eslint-disable-next-line no-undef
      global
    : undefined
}

const g = getGlobal() as any

// the add-in command functions need to be available in global scope
g.encrypt = encrypt
