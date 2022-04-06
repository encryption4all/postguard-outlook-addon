/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

import { ComposeMail } from '@e4a/irmaseal-mail-utils/dist/index'
import { createMimeMessage } from 'mimetext'
import {
  storeMailAsPlainLocally,
  IAttachmentContent,
  htmlBodyType,
  getItemRestId
} from '../helpers/utils'

// eslint-disable-next-line no-undef
var Buffer = require('buffer/').Buffer

var mailboxItem: Office.MessageCompose
var globalEvent
var isEncryptMode: boolean = false

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

  isEncryptMode = true
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
      let returnBody: string
      const body: string = asyncResult.value
      const pattern = /<body[^>]*>((.|[\n\r])*)<\/body>/im
      const arrayMatches = pattern.exec(body)
      if (arrayMatches === null) {
        returnBody = body
      } else {
        const mailBody = arrayMatches[1]
        returnBody = mailBody
      }
      if (returnBody !== '') resolve(returnBody)
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
              isInline: attachment.isInline,
              id: attachment.id
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
  let pk
  const response = await fetch(`${hostname}/v2/parameters`)

  // if response is not ok, try to get PK from localStorage
  if (!response.ok || (response.status < 200 && response.status > 299)) {
    const cachedPK = window.localStorage.getItem('pk')
    if (cachedPK) {
      pk = JSON.parse(cachedPK)
    } else {
      // if PK also not available in localStorage remove it
      const errorMsg = `Cannot retrieve publickey from ${hostname} and not stored within localStorage`
      showInfoMessage(errorMsg)
      return Promise.reject(errorMsg)
    }
  }

  pk = await response.json()
  window.localStorage.setItem('pk', JSON.stringify(pk))
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
  let isAlreadyEncrypted: boolean = false

  if (attachments !== undefined) {
    let usePostGuard = false
    for (let i = 0; i < attachments.length; i++) {
      const attachment = attachments[i]
      isAlreadyEncrypted = attachment.filename === 'postguard.encrypted'

      if (!attachment.isInline) {
        hasAttachments = true
        if (!usePostGuard) {
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

  if (!isAlreadyEncrypted) {
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

    Office.context.mailbox.item.subject.setAsync(
      'PostGuard encrypted email',
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(
            'Changing subject failed with error: ' + asyncResult.error.message
          )
        } else {
          mailboxItem.body.setAsync(
            '<b>This is a PostGuard encrypted email</b>',
            { coercionType: Office.CoercionType.Html },
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(
                  'Changing body failed with error: ' +
                    asyncResult.error.message
                )
              } else {
                const b64 = Buffer.from(ct).toString('base64')
                mailboxItem.addFileAttachmentFromBase64Async(
                  b64,
                  'postguard.encrypted',
                  (asyncResult) => {
                    if (
                      asyncResult.status === Office.AsyncResultStatus.Failed
                    ) {
                      console.log(
                        'Adding attachment failed with error: ' +
                          asyncResult.error.message
                      )
                    } else {
                      storeMailAsPlainLocally(
                        token,
                        jsonInnerMail,
                        attachments,
                        'PostGuard'
                      )

                      showInfoMessage(
                        'Successfully encrypted email, press Send to send the email'
                      )

                      if (attachments !== undefined) {
                        for (let i = 0; i < attachments.length; i++) {
                          const attachment = attachments[i]
                          mailboxItem.removeAttachmentAsync(
                            attachment.id,
                            (asyncResult) => {
                              if (
                                asyncResult.status ===
                                Office.AsyncResultStatus.Failed
                              ) {
                                console.log(
                                  'Changing subject failed with error: ' +
                                    asyncResult.error.message
                                )
                              }
                            }
                          )
                        }
                      }
                    }
                  }
                )
              }
            }
          )
        }
      }
    )
  } else {
    showInfoMessage(
      'Mail is already encrypted with PostGuard, cannot encrypt again. Please send the email.'
    )
  }
}

function showInfoMessage(msg: string) {
  const msgDetails: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: msg,
    icon: 'Icon.80x80',
    persistent: true
  }

  console.log('Info msg: ', msg)

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    'action',
    msgDetails
  )
  globalEvent.completed()
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
    console.log('Trying to close login dialog')
    loginDialog.close()
    console.log('Valid token: ', JSON.stringify(messageFromDialog.result))
    console.log('Logginger: ', JSON.stringify(messageFromDialog.logging))
    if (isEncryptMode) {
      encryptAndsendMail(messageFromDialog.result.accessToken)
    } else {
      showDecryptPopup(messageFromDialog.result.accessToken)
    }
  } else {
    // Something went wrong with authentication or the authorization of the web application.
    loginDialog.close()
    console.log('Error: ', JSON.stringify(messageFromDialog.error.toString()))
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
g.decrypt = decrypt

// DECRYPTION TEST IN DIALOG
// eslint-disable-next-line no-unused-vars
var decryptDialog: Office.Dialog

function decrypt(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: 'Decrypting email via PostGuard',
    icon: 'Icon.80x80',
    persistent: true
  }

  globalEvent = event

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    'action',
    message
  )

  if (isIrmasealEmail()) {
    showLoginPopup('/fallbackauthdialog.html')
  } else {
    globalEvent.completed()
    const message: Office.NotificationMessageDetails = {
      type: Office.MailboxEnums.ItemNotificationMessageType
        .InformationalMessage,
      message: 'Not a PostGuard email, cannot decrypt.',
      icon: 'Icon.80x80',
      persistent: true
    }
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      'action',
      message
    )
  }
}

// first attachment name must be 'postguard.encrypted', and subject must be 'PostGuard encrypted email'
function isIrmasealEmail() {
  const item = Office.context.mailbox.item
  if (item.attachments.length != 0) {
    const attachmentName = item.attachments[0].name
    const subjectTitle = item.subject
    if (
      attachmentName === 'postguard.encrypted' &&
      subjectTitle === 'PostGuard encrypted email'
    ) {
      console.log('It is a PostGuard email!')
      return true
    } else {
      console.log('No PostGuard email')
      return false
    }
  } else {
    console.log('No PostGuard email')
    return false
  }
}

function showDecryptPopup(token) {
  const b64 = Buffer.from(token).toString('base64')
  const fullUrl =
    location.protocol +
    '//' +
    location.hostname +
    (location.port ? ':' + location.port : '') +
    '/decrypt.html' +
    '?token=' +
    b64 +
    '&mailid=' +
    getItemRestId() +
    '&recipient=' +
    Office.context.mailbox.userProfile.emailAddress +
    '&attachmentid=' +
    Office.context.mailbox.item.attachments[0].id

  // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
  Office.context.ui.displayDialogAsync(
    fullUrl,
    { height: 50, width: 10 },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        if (result.error.code === 12007) {
          showDecryptPopup(token) // Recursive call
        } else {
          console.log('Other error: ', result.error)
        }
      } else {
        console.log('Decryptdialog has initialized, ', result)
        decryptDialog = result.value
        decryptDialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          processDecryptMessage
        )
      }
    }
  )
}

function processDecryptMessage(msg) {
  globalEvent.completed()
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: msg.message,
    icon: 'Icon.80x80',
    persistent: true
  }
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    'action',
    message
  )
}
