/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, Office */

import 'web-streams-polyfill'

import { ComposeMail } from '@e4a/irmaseal-mail-utils/dist/index'
import { createMimeMessage } from 'mimetext'
import {
  storeMailAsPlainLocally,
  IAttachmentContent,
  htmlBodyType,
  getItemRestId,
  isPostGuardEmail,
  newReadableStreamFromArray,
  getGlobal,
  getPostGuardHeaders
} from '../helpers/utils'
import type { Policy } from 'attribute-form/AttributeForm/AttributeForm.svelte'

// eslint-disable-next-line no-undef
var Buffer = require('buffer/').Buffer

var mailboxItem: Office.MessageCompose
var globalEvent
var isEncryptMode: boolean = false
var isExtendedEncryption: boolean = false

const hostname = 'https://stable.irmaseal-pkg.ihub.ru.nl'
const email_attribute = 'pbdf.sidn-pbdf.email.email'

const mod_promise = import('@e4a/irmaseal-wasm-bindings')

import * as getLogger from 'webpack-log'
const encryptLog = getLogger({ name: 'PostGuard encrypt log' })
const decryptLog = getLogger({ name: 'PostGuard decrypt log' })

/**
 * The initialization function
 */
Office.initialize = () => {
  Office.onReady(() => {
    mailboxItem = Office.context.mailbox.item
    delete window.alert // assures alert works
    delete window.confirm // assures confirm works
    delete window.prompt // assures prompt works
  })
}

/**
 * Entry point function for encryption
 * @param event The AddinCommands Event
 */
// eslint-disable-next-line no-unused-vars
function encrypt(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: 'Encrypting email with PostGuard',
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

/**
 * Entry point function for encryption
 * @param event The AddinCommands Event
 */
// eslint-disable-next-line no-unused-vars
function encryptExtended(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: 'Encrypting email with PostGuard',
    icon: 'Icon.80x80',
    persistent: true
  }

  globalEvent = event

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    'action',
    message
  )

  isEncryptMode = true
  isExtendedEncryption = true
  showLoginPopup('/fallbackauthdialog.html')
}

/**
 * Gets the To of the current mail item
 * @returns The To as an array of string
 */
function getRecipientEmails(): Promise<string[]> {
  return new Promise(function (resolve) {
    mailboxItem.to.getAsync((recipients) => {
      let recipientMails = new Array()
      recipients.value.forEach((recipient) => {
        recipientMails.push(recipient.emailAddress.toLowerCase())
      })
      resolve(recipientMails)
    })
  })
}

/**
 * Gets the Ccs of the current mail item
 * @returns The Ccs as an array of string
 */
function getCcRecipientEmails(): Promise<string[]> {
  return new Promise(function (resolve) {
    mailboxItem.cc.getAsync((recipients) => {
      let recipientMails = new Array()
      recipients.value.forEach((recipient) => {
        recipientMails.push(recipient.emailAddress.toLowerCase())
      })
      resolve(recipientMails)
    })
  })
}

/**
 * Gets the Bccs of the current mail item
 * @returns The Bccs as an array of string
 */
function getBccRecipientEmails(): Promise<string[]> {
  return new Promise(function (resolve) {
    mailboxItem.bcc.getAsync((recipients) => {
      let recipientMails = new Array()
      recipients.value.forEach((recipient) => {
        recipientMails.push(recipient.emailAddress.toLowerCase())
      })
      resolve(recipientMails)
    })
  })
}

/**
 * Gets the body of the current mail item
 * @returns The body
 */
async function getMailBody(): Promise<string> {
  return new Promise(function (resolve, reject) {
    mailboxItem.body.getAsync(Office.CoercionType.Html, (asyncResult) => {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        reject('Body async failed')
      } else {
        let returnBody: string = asyncResult.value
        if (returnBody !== '') resolve(returnBody)
        else reject('Please add text to the body in the email')
      }
    })
  })
}

/**
 * Gets the subject from the current mail item
 * @returns The subject
 */
async function getMailSubject(): Promise<string> {
  return new Promise(function (resolve, reject) {
    mailboxItem.subject.getAsync((asyncResult) => {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        reject('Subject async failed')
      } else {
        const subject: string = asyncResult.value
        if (subject !== '') resolve(subject)
        else reject('Please add a subject to the email')
      }
    })
  })
}

/**
 * Collects all metadata and content of attachment belonging to the current mail item
 * @returns Array of IAttachmentContent objects
 */
async function getMailAttachments(): Promise<IAttachmentContent[]> {
  return new Promise(function (resolve, reject) {
    mailboxItem.getAttachmentsAsync(async (asyncResult) => {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        reject('Attachments async failed')
      } else {
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
      }
    })
  })
}

/**
 * Returns promise containing content of the attachment
 * @param attachmentId ID of the attachment
 * @returns Promise containing content of the attachment
 */
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

/**
 * Gets public key from PKG, or local storage if PKG is not available
 * @returns The public key
 */
async function getPublicKey(): Promise<any> {
  let response
  let headers = {
    'X-Postguard-Client-Version': `Outlook,${Office.context.diagnostics.version},pg4ol,0.0.1`
  }
  console.log(`Headers: ${headers}`)
  try {
    response = await fetch(`${hostname}/v2/parameters`, {
      headers: { 'X-Postguard-Client-Version': getPostGuardHeaders() }
    })
  } catch (e) {
    encryptLog.error(e)
  }
  let pk

  // if response is not ok, try to get PK from localStorage
  if (!response.ok || (response.status < 200 && response.status > 299)) {
    const cachedPK = window.localStorage.getItem('pk')
    if (cachedPK) {
      pk = JSON.parse(cachedPK)
    } else {
      throw 'Cannot retrieve public key'
    }
  } else {
    pk = await response.json()
    window.localStorage.setItem('pk', JSON.stringify(pk))
  }
  return pk
}

/**
 * Encrypts and sends the email
 * @param token The authentication token for the Graph API
 */
async function encryptAndSendEmail(token, policy: Policy = null) {
  const pk = await getPublicKey()

  const [mod] = await Promise.all([mod_promise])

  const sender = Office.context.mailbox.userProfile.emailAddress

  const timestamp = Math.round(Date.now() / 1000)

  const recipientEmails: string[] = await getRecipientEmails()
  const ccRecipientEmails: string[] = await getCcRecipientEmails()
  const bccRecipientEmails: string[] = await getBccRecipientEmails()

  // if BCC recipients available , abort operations
  if (bccRecipientEmails.length > 0) {
    bccMsgAndDialog()
    return
  }

  const allRecipientsCount = recipientEmails.length + ccRecipientEmails.length

  if (allRecipientsCount == 0) {
    throw 'Please add recipients to the email!'
  }

  let allPolicies = {}

  if (policy === null) {
    const policies = recipientEmails.reduce((total, recipient) => {
      total[recipient] = {
        ts: timestamp,
        con: [{ t: email_attribute, v: recipient }]
      }
      return total
    }, {})

    const ccPolicies = ccRecipientEmails.reduce((total, recipient) => {
      total[recipient] = {
        ts: timestamp,
        con: [{ t: email_attribute, v: recipient }]
      }
      return total
    }, {})

    allPolicies = { ...policies, ...ccPolicies }
  } else {
    for (var id in policy) {
      allPolicies[id] = {
        ts: timestamp,
        con: policy[id]
      }
    }
  }

  encryptLog.info('Encrypting using the following policies: ', allPolicies)

  let mailBody = await getMailBody()

  const mailSubject = await getMailSubject()
  encryptLog.info('Mail subject: ', mailSubject)

  let attachments: IAttachmentContent[] = await getMailAttachments()

  // Use createMimeMessage to create inner MIME mail
  const msg = createMimeMessage()
  msg.setSender(sender)
  msg.setSubject(mailSubject)

  recipientEmails.length > 0 && msg.setRecipient(recipientEmails)
  ccRecipientEmails.length > 0 && msg.setCc(ccRecipientEmails)
  bccRecipientEmails.length > 0 && msg.setBcc(bccRecipientEmails)

  // ComposeMail only used for outer mail
  const composeMail = new ComposeMail()
  composeMail.setSubject('PostGuard Encrypted Email')
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
    for (let i = 0; i < attachments.length; i++) {
      const attachment = attachments[i]

      if (!attachment.isInline) {
        hasAttachments = true
        const input = new TextEncoder().encode(attachment.content)
        encryptLog.info('Attachment bytes length: ', input.byteLength)
        msg.setAttachment(
          attachment.filename,
          'application/octet-stream',
          attachment.content
        )
      } else {
        // replace inline image in body
        const imageContentIDToReplace = `cid:${attachment.filename}@.*"`
        const regex = new RegExp(imageContentIDToReplace, 'g')
        mailBody = mailBody.replace(
          regex,
          `data:image;base64,${attachment.content}"`
        )
      }
    }
  }

  encryptLog.info('Mailbody: ', mailBody)
  msg.setMessage('text/html', mailBody)

  // encrypt inner MIME mail
  const innerMail = msg.asRaw()
  const plainBytes: Uint8Array = new TextEncoder().encode(innerMail)
  const readable = newReadableStreamFromArray(plainBytes)
  let ct = new Uint8Array(0)
  const writable = new WritableStream({
    write(chunk) {
      ct = new Uint8Array([...ct, ...chunk])
    }
  })

  await mod.seal(pk.publicKey, allPolicies, readable, writable)

  // get outer mail to send email via Graph API
  composeMail.setPayload(ct)
  const outerMail = composeMail.getMimeMail()
  const message = Buffer.from(outerMail).toString('base64')
  const sendMessageUrl = 'https://graph.microsoft.com/v1.0/me/sendMail'
  encryptLog.info('Trying to send email via ', sendMessageUrl)

  $.ajax({
    type: 'POST',
    contentType: 'text/plain',
    url: sendMessageUrl,
    data: message,
    headers: {
      Authorization: 'Bearer ' + token
    },
    success: function () {
      encryptLog.info(
        'Sendmail success, now trying to store mail locally, and clear mail'
      )

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

      storeMailAsPlainLocally(
        token,
        jsonInnerMail,
        attachments,
        'PostGuard Sent'
      )

      clearCurrentEmail(attachments)

      successMsgAndDialog()
    }
  }).fail(function ($xhr) {
    var data = $xhr.responseJSON
    encryptLog.error('Ajax error send mail: ', data)
    throw 'Error when sending mail, please try again or contact administrator'
  })
}

/**
 * Adds info message to email and shows dialog that PostGuard does not support BCC yet
 */
function bccMsgAndDialog() {
  showInfoMessage(
    'PostGuard does not support BCCs. Please remove BCCs, or send the mail unencrypted.'
  )
  /*var fullUrl =
    'https://' +
    location.hostname +
    (location.port ? ':' + location.port : '') +
    '/bcc.html'

  Office.context.ui.displayDialogAsync(
    fullUrl,
    { height: 20, width: 30 },
    // eslint-disable-next-line no-unused-vars
    function (result) {
      encryptLog.info('Bccdialog has initialized.')
    }
  )*/
}

/**
 * Adds info message to email and shows dialog that message is successfully encrypted and send
 */
function successMsgAndDialog() {
  showInfoMessage('Successfully encrypted and sent')
  var fullUrl =
    'https://' +
    location.hostname +
    (location.port ? ':' + location.port : '') +
    '/success.html'

  Office.context.ui.displayDialogAsync(
    fullUrl,
    { height: 40, width: 30 },
    // eslint-disable-next-line no-unused-vars
    function (result) {
      encryptLog.info('Successdialog has initialized.')
    }
  )
}

/**
 * Clears the current mail, which means that to, cc, bcc, subject, body and attachments are removed
 * @param attachments The attachments of the mail
 */

function clearCurrentEmail(attachments) {
  mailboxItem.to.setAsync([], function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      encryptLog.error(
        'Clearing subject failed with error: ' + asyncResult.error.message
      )
    }
  })
  mailboxItem.cc.setAsync([], function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      encryptLog.error(
        'Clearing subject failed with error: ' + asyncResult.error.message
      )
    }
  })
  mailboxItem.bcc.setAsync([], function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      encryptLog.error(
        'Clearing subject failed with error: ' + asyncResult.error.message
      )
    }
  })
  mailboxItem.subject.setAsync('', function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      encryptLog.error(
        'Clearing subject failed with error: ' + asyncResult.error.message
      )
    }
  })
  mailboxItem.body.setAsync('', (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      encryptLog.error(
        'Clearing body failed with error: ' + asyncResult.error.message
      )
    }
  })
  if (attachments !== undefined) {
    for (let i = 0; i < attachments.length; i++) {
      const attachment = attachments[i]
      mailboxItem.removeAttachmentAsync(attachment.id, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          encryptLog.error(
            'Clearing attachment failed with error: ' +
              asyncResult.error.message
          )
        }
      })
    }
  }
}

/**
 * Displays a message
 * @param msg The message to be displayes
 */
function showInfoMessage(msg: string) {
  const msgDetails: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: msg,
    icon: 'Icon.80x80',
    persistent: true
  }
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    'action',
    msgDetails
  )
  globalEvent.completed()
}

var attributeDialog

/**
 * Shows dialog to add attributes for each recipient
 * @param accessToken Token to access graph API
 */
async function addAttributes(accessToken: string) {
  const recipientEmails: string[] = await getRecipientEmails()
  const ccRecipientEmails: string[] = await getCcRecipientEmails()
  const bccRecipientEmails: string[] = await getBccRecipientEmails()

  // if BCC recipients available , abort operations
  if (bccRecipientEmails.length > 0) {
    bccMsgAndDialog()
    return
  }

  const allRecipientsCount = recipientEmails.length + ccRecipientEmails.length

  if (allRecipientsCount == 0) {
    throw 'Please add recipients to the email!'
  }

  const recipientsStringified = JSON.stringify(
    recipientEmails.concat(ccRecipientEmails)
  )

  const b64Recipients = Buffer.from(recipientsStringified).toString('base64')
  const b64Token = Buffer.from(accessToken).toString('base64')

  var fullUrl =
    location.protocol +
    '//' +
    location.hostname +
    (location.port ? ':' + location.port : '') +
    '/attributes.html' +
    '?recipients=' +
    b64Recipients +
    '&token=' +
    b64Token

  Office.context.ui.displayDialogAsync(
    fullUrl,
    { height: 40, width: 20 },
    function (result) {
      encryptLog.info('Attributedialog has initialized.')
      attributeDialog = result.value
      attributeDialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        processAttributesMessage
      )
    }
  )
}

/**
 *  This handler responds to the success or failure message that the pop-up dialog receives from the identity provider and access token provider.
 * @param arg The arg object passed from the dialog
 */

async function processAttributesMessage(arg) {
  let messageFromDialog = JSON.parse(arg.message)
  attributeDialog.close()

  if (messageFromDialog.status === 'success') {
    const policies = messageFromDialog.result.policy
    const recipients = Object.keys(policies)

    Office.context.mailbox.item.to.setAsync(recipients, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        encryptAndSendEmail(
          messageFromDialog.result.accessToken,
          messageFromDialog.result.policy
        ).catch((err) => {
          encryptLog.error(err)
          showInfoMessage(err)
        })
      } else {
        throw 'Could not set recipients in compose window'
      }
    })
  } else {
    encryptLog.error(
      'Error: ',
      JSON.stringify(messageFromDialog.error.toString())
    )
  }
}

var loginDialog

/**
 *  This handler responds to the success or failure message that the pop-up dialog receives from the identity provider and access token provider.
 * @param arg The arg object passed from the dialog
 */
async function processMessage(arg) {
  let messageFromDialog = JSON.parse(arg.message)
  encryptLog.info(`After auth: ${JSON.stringify(messageFromDialog)}`)

  if (messageFromDialog.status === 'success') {
    loginDialog.close()
    if (isEncryptMode) {
      g.msgFunc = showInfoMessage
      if (isExtendedEncryption) {
        addAttributes(messageFromDialog.result.accessToken).catch((err) => {
          encryptLog.error(err)
          showInfoMessage(err)
        })
      } else {
        encryptAndSendEmail(messageFromDialog.result.accessToken).catch(
          (err) => {
            encryptLog.error(err)
            showInfoMessage(err)
          }
        )
      }
    } else {
      showDecryptPopup(messageFromDialog.result.accessToken)
    }
  } else {
    // Something went wrong with authentication or the authorization of the web application... try again
    encryptLog.error('Error: ', JSON.stringify(messageFromDialog))
  }
}

/**
 * Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
 * @param url The HTML url pointing to the login popup
 */

function showLoginPopup(url: string) {
  var fullUrl =
    location.protocol +
    '//' +
    location.hostname +
    (location.port ? ':' + location.port : '') +
    url +
    '?currentAccountMail=' +
    Office.context.mailbox.userProfile.emailAddress

  Office.context.ui.displayDialogAsync(
    fullUrl,
    { height: 60, width: 30 },
    function (result) {
      encryptLog.info('Logindialog has initialized.')
      loginDialog = result.value
      loginDialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        processMessage
      )
    }
  )
}

// eslint-disable-next-line no-unused-vars
var decryptDialog: Office.Dialog

/**
 * Decrypts current message via popup
 * @param event The AddinCommands Event
 */
function decrypt(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: 'Decrypting email with PostGuard',
    icon: 'Icon.80x80',
    persistent: true
  }

  globalEvent = event

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    'action',
    message
  )

  if (isPostGuardEmail()) {
    showLoginPopup('/fallbackauthdialog.html')
  } else {
    showInfoMessage('This is not a PostGuard Email, cannot decrypt.')
  }
}

/**
 * Displays the decrypt popup
 * @param token Authentication token for Graph API
 */
function showDecryptPopup(token: string) {
  const b64 = Buffer.from(token).toString('base64')

  const fullUrl =
    'https://' +
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
      .replace('/', '-')
      .replace('+', '_') +
    '&sender=' +
    Office.context.mailbox.item.sender.emailAddress

  // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
  Office.context.ui.displayDialogAsync(
    fullUrl,
    { height: 80, width: 30 },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        if (result.error.code === 12007) {
          showDecryptPopup(token) // Recursive call
        } else {
          decryptLog.error('Decryptdialog error: ', result.error)
          showInfoMessage(result.error.message)
        }
      } else {
        decryptLog.info('Decryptdialog has initialized, ', result)
        decryptDialog = result.value
        decryptDialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          processDecryptMessage
        )
      }
    }
  )
}

/**
 * Event handler function receiving message form decrypt dialog
 * Passing message on to showInfoMessage function
 * @param arg object passed from decrypt dialog.
 */
function processDecryptMessage(arg) {
  showInfoMessage(arg.message)
}

// the add-in command functions need to be available in global scope
const g = getGlobal() as any
g.encrypt = encrypt
g.encryptExt = encryptExtended
g.decrypt = decrypt
